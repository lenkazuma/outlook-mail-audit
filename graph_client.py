"""Microsoft Graph client with a GET-only mailbox API surface."""
from __future__ import annotations
import os
import re
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = ["User.Read", "Mail.Read", "MailboxSettings.Read"]
RESERVED_SCOPES = {"openid", "profile", "offline_access"}


def normalize_scopes(scopes: list[str] | None) -> list[str]:
    # MSAL adds OpenID scopes internally; older .env files may include them.
    result = [scope for scope in (scopes or DEFAULT_SCOPES) if scope.lower() not in RESERVED_SCOPES]
    allowed = {scope.lower() for scope in DEFAULT_SCOPES}
    if not result or any(scope.lower() not in allowed for scope in result):
        raise ValueError("Only User.Read, Mail.Read and MailboxSettings.Read scopes are supported.")
    return list(dict.fromkeys(result))


def graph_url(path_or_url: str) -> str:
    """Validate pagination destinations before attaching an access token."""
    if path_or_url.startswith(("http://", "https://")):
        url = path_or_url
    else:
        if path_or_url.startswith("//") or ":" in path_or_url.split("/", 1)[0]:
            raise ValueError("Invalid Graph path.")
        url = f"{GRAPH_BASE}/{path_or_url.lstrip('/')}"
    parts = urlsplit(url)
    if (parts.scheme != "https" or parts.netloc.lower() != "graph.microsoft.com"
            or not parts.path.startswith("/v1.0/") or parts.fragment):
        raise ValueError("Only HTTPS Microsoft Graph v1.0 URLs are allowed.")
    return url


class GraphClient:
    def __init__(self, client_id: str, tenant: str = "common", scopes: list[str] | None = None,
                 cache_path: str = ".token_cache.bin") -> None:
        self.scopes = normalize_scopes(scopes)
        if not re.fullmatch(r"[A-Za-z0-9.-]+", tenant):
            raise ValueError("Invalid tenant identifier.")
        try:
            import msal
            import requests
        except ImportError as exc:
            raise RuntimeError("Install Graph dependencies: python -m pip install -r requirements.txt") from exc
        self._requests = requests
        self.cache_path = Path(cache_path)
        self.token_cache = msal.SerializableTokenCache()
        if self.cache_path.exists():
            self.token_cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
        self.app = msal.PublicClientApplication(
            client_id=client_id, authority=f"https://login.microsoftonline.com/{tenant}",
            token_cache=self.token_cache,
        )
        self._access_token: str | None = None

    def _persist_cache(self) -> None:
        if self.token_cache.has_state_changed:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            self.cache_path.write_text(self.token_cache.serialize(), encoding="utf-8")
            self.cache_path.chmod(0o600)  # Windows: also use user-only ACLs for the folder.

    def login(self) -> None:
        accounts = self.app.get_accounts()
        if len(accounts) > 1:
            raise RuntimeError("Multiple cached accounts found. Use a separate cache per account.")
        result = self.app.acquire_token_silent(self.scopes, account=accounts[0]) if accounts else None
        if not result:
            flow = self.app.initiate_device_flow(scopes=self.scopes)
            if "user_code" not in flow:
                raise RuntimeError(f"Device login unavailable: {flow.get('error', 'unknown error')}")
            print(flow["message"])
            result = self.app.acquire_token_by_device_flow(flow)
        if not result or "access_token" not in result:
            raise RuntimeError(f"Token acquisition failed: {(result or {}).get('error', 'unknown error')}")
        self._access_token = result["access_token"]
        self._persist_cache()

    def get(self, path_or_url: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        url = graph_url(path_or_url)
        if not self._access_token:
            raise RuntimeError("Not authenticated. Call login() first.")
        response = self._requests.get(
            url, headers={"Authorization": f"Bearer {self._access_token}", "Accept": "application/json"},
            params=params, timeout=60, allow_redirects=False,
        )
        if not 200 <= response.status_code < 300:
            # Do not include message URLs or server response bodies in error logs.
            raise RuntimeError(f"Microsoft Graph GET failed (HTTP {response.status_code}).")
        return response.json()

    def get_all_pages(self, path_or_url: str, params: dict[str, Any] | None = None) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        next_url: str | None = path_or_url
        visited: set[str] = set()
        while next_url:
            if next_url in visited:
                raise RuntimeError("Microsoft Graph returned a repeated pagination link.")
            visited.add(next_url)
            payload = self.get(next_url, params=params)
            items.extend(payload.get("value", []))
            next_url = payload.get("@odata.nextLink")
            params = None
        return items


def load_env(path: str = ".env") -> None:
    env_path = Path(path)
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        os.environ.setdefault(key.strip(), value)
