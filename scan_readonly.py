from __future__ import annotations

import argparse
import json
import os
from collections import Counter
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

from graph_client import GraphClient, load_env
from audit_utils import positive_int, write_csv


def safe_get(d: Dict[str, Any], path: list[str], default: Any = "") -> Any:
    cur: Any = d
    for key in path:
        if not isinstance(cur, dict):
            return default
        cur = cur.get(key)
    return cur if cur is not None else default


def flatten_rule(rule: Dict[str, Any]) -> Dict[str, Any]:
    conditions = rule.get("conditions") or {}
    actions = rule.get("actions") or {}
    sender_contains = conditions.get("senderContains", []) or []
    body_contains = conditions.get("bodyContains", []) or []
    subject_contains = conditions.get("subjectContains", []) or []
    move_to_folder = safe_get(actions, ["moveToFolder"], "")
    return {
        "id": rule.get("id", ""),
        "displayName": rule.get("displayName", ""),
        "isEnabled": rule.get("isEnabled", False),
        "hasError": rule.get("hasError", False),
        "isReadOnly": rule.get("isReadOnly", False),
        "sequence": rule.get("sequence", ""),
        "stopProcessingRules": actions.get("stopProcessingRules", False),
        "moveToFolder": move_to_folder,
        "senderContains": "; ".join(sender_contains),
        "subjectContains": "; ".join(subject_contains),
        "bodyContains": "; ".join(body_contains),
    }


def to_dt(dt_str: str) -> datetime:
    return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))


def traverse_folders(client: GraphClient) -> List[Dict[str, Any]]:
    result: List[Dict[str, Any]] = []
    root_folders = client.get_all_pages(
        "/me/mailFolders",
        params={
            "$top": 200,
            "$select": "id,displayName,parentFolderId,totalItemCount,unreadItemCount,childFolderCount,isHidden",
            "includeHiddenFolders": "true",
        },
    )

    queue = list(root_folders)
    seen = set()
    while queue:
        folder = queue.pop(0)
        folder_id = folder.get("id")
        if not folder_id or folder_id in seen:
            continue
        seen.add(folder_id)
        result.append(folder)
        child_count = folder.get("childFolderCount", 0) or 0
        if child_count > 0:
            children = client.get_all_pages(
                f"/me/mailFolders/{folder_id}/childFolders",
                params={
                    "$top": 200,
                    "$select": "id,displayName,parentFolderId,totalItemCount,unreadItemCount,childFolderCount,isHidden",
                    "includeHiddenFolders": "true",
                },
            )
            queue.extend(children)
    return result


def get_inbox_id(client: GraphClient) -> str:
    # The well-known identifier works regardless of the mailbox display language.
    return client.get("/me/mailFolders/inbox", params={"$select": "id"})["id"]


def collect_messages(
    client: GraphClient,
    folder_id: str,
    folder_name: str,
    cutoff: datetime,
    max_messages: int,
) -> List[Dict[str, Any]]:
    params = {
        "$top": 50,
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,receivedDateTime,parentFolderId,from,sender,conversationId,categories,isRead,importance,hasAttachments,webLink,internetMessageId",
    }
    path = f"/me/mailFolders/{folder_id}/messages"
    rows: List[Dict[str, Any]] = []
    pages = 0
    visited = set()
    while path and len(rows) < max_messages:
        if path in visited:
            raise RuntimeError("Repeated message pagination link.")
        visited.add(path)
        payload = client.get(path, params=params if pages == 0 else None)
        pages += 1
        items = payload.get("value", [])
        if not items:
            break
        stop_for_cutoff = False
        for m in items:
            rdt = m.get("receivedDateTime")
            if not rdt:
                continue
            received = to_dt(rdt)
            if received < cutoff:
                stop_for_cutoff = True
                break
            from_addr = safe_get(m, ["from", "emailAddress", "address"], "")
            from_name = safe_get(m, ["from", "emailAddress", "name"], "")
            rows.append(
                {
                    "id": m.get("id", ""),
                    "folderId": folder_id,
                    "folderName": folder_name,
                    "receivedDateTime": rdt,
                    "subject": (m.get("subject") or "").replace("\n", " ").strip(),
                    "fromAddress": from_addr.lower(),
                    "fromName": from_name,
                    "categories": "; ".join(m.get("categories", []) or []),
                    "isRead": m.get("isRead", False),
                    "importance": m.get("importance", ""),
                    "hasAttachments": m.get("hasAttachments", False),
                    "webLink": m.get("webLink", ""),
                    "internetMessageId": m.get("internetMessageId", ""),
                    "conversationId": m.get("conversationId", ""),
                }
            )
            if len(rows) >= max_messages:
                break
        if stop_for_cutoff:
            break
        path = payload.get("@odata.nextLink")
        params = None
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description="Read-only Outlook mailbox scanner")
    parser.add_argument("--days", type=positive_int, default=365, help="How many days of messages to scan")
    parser.add_argument(
        "--max-messages-per-folder",
        type=positive_int,
        default=5000,
        help="Safety cap for messages collected per folder",
    )
    parser.add_argument(
        "--suggest-threshold",
        type=positive_int,
        default=5,
        help="Minimum unsorted message count before suggesting sender-based rule",
    )
    parser.add_argument("--output-dir", default="output", help="Output directory path")
    parser.add_argument("--token-cache", default=".token_cache.bin", help="Private token cache path")
    args = parser.parse_args()

    load_env()
    client_id = os.getenv("GRAPH_CLIENT_ID", "").strip()
    tenant = os.getenv("GRAPH_TENANT", "common").strip() or "common"
    scopes = os.getenv("GRAPH_SCOPES", "").strip()
    scope_list = scopes.split() if scopes else None
    if not client_id or client_id == "your_app_client_id":
        raise RuntimeError(
            "GRAPH_CLIENT_ID is missing. Put it in .env or env vars. See .env.example."
        )

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    cutoff = datetime.now(UTC) - timedelta(days=args.days)

    client = GraphClient(client_id=client_id, tenant=tenant, scopes=scope_list, cache_path=args.token_cache)
    client.login()

    me = client.get("/me?$select=id,displayName,userPrincipalName,mail")
    folders = traverse_folders(client)
    inbox_id = get_inbox_id(client)
    rules = client.get_all_pages(f"/me/mailFolders/{inbox_id}/messageRules", params={"$top": 200})

    folder_rows = [
        {
            "id": f.get("id", ""),
            "displayName": f.get("displayName", ""),
            "parentFolderId": f.get("parentFolderId", ""),
            "totalItemCount": f.get("totalItemCount", 0),
            "unreadItemCount": f.get("unreadItemCount", 0),
            "childFolderCount": f.get("childFolderCount", 0),
            "isHidden": f.get("isHidden", False),
        }
        for f in folders
    ]
    folder_rows.sort(key=lambda x: x["displayName"].lower())

    rule_rows = [flatten_rule(r) for r in rules]
    enabled_sender_contains = set()
    for r in rule_rows:
        if not r["isEnabled"]:
            continue
        senders = [s.strip().lower() for s in str(r["senderContains"]).split(";") if s.strip()]
        enabled_sender_contains.update(senders)

    all_messages: List[Dict[str, Any]] = []
    for f in folders:
        folder_id = f.get("id", "")
        folder_name = f.get("displayName", "")
        if not folder_id:
            continue
        msgs = collect_messages(
            client=client,
            folder_id=folder_id,
            folder_name=folder_name,
            cutoff=cutoff,
            max_messages=args.max_messages_per_folder,
        )
        all_messages.extend(msgs)

    unsorted = [
        m
        for m in all_messages
        if m["folderId"] == inbox_id and not m.get("categories")
    ]
    sender_counter = Counter([m["fromAddress"] for m in unsorted if m["fromAddress"]])

    sender_rows = []
    for sender, count in sender_counter.most_common():
        if sender in enabled_sender_contains:
            already_in_rules = True
        else:
            already_in_rules = any(token and token in sender for token in enabled_sender_contains)
        sender_rows.append(
            {
                "sender": sender,
                "unsortedCount": count,
                "alreadyCoveredByExistingSenderRule": already_in_rules,
            }
        )

    proposed_rules = []
    for row in sender_rows:
        if row["unsortedCount"] < args.suggest_threshold:
            continue
        if row["alreadyCoveredByExistingSenderRule"]:
            continue
        sender = row["sender"]
        domain = sender.split("@")[-1] if "@" in sender else "unknown-domain"
        proposed_rules.append(
            {
                "type": "senderContains",
                "match": sender,
                "reason": f"Inbox unsorted {row['unsortedCount']} messages in last {args.days} days",
                "suggestedFolder": f"Auto/{domain}",
            }
        )

    (out_dir / "mailbox_profile.json").write_text(
        json.dumps(
            {
                "scannedAtUtc": datetime.now(UTC).isoformat(),
                "account": {
                    "displayName": me.get("displayName"),
                    "userPrincipalName": me.get("userPrincipalName"),
                    "mail": me.get("mail"),
                },
                "days": args.days,
                "foldersScanned": len(folders),
                "messagesScanned": len(all_messages),
                "inboxUnsortedCount": len(unsorted),
                "rulesCount": len(rule_rows),
                "proposedRuleCount": len(proposed_rules),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    (out_dir / "folders.json").write_text(json.dumps(folders, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "rules.json").write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "proposed_rules.json").write_text(
        json.dumps(proposed_rules, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    write_csv(out_dir / "folders.csv", folder_rows)
    write_csv(out_dir / "rules_inventory.csv", rule_rows)
    write_csv(out_dir / "all_messages_metadata.csv", all_messages)
    write_csv(out_dir / "unsorted_inbox_messages.csv", unsorted)
    write_csv(out_dir / "unsorted_senders.csv", sender_rows)

    print("Read-only scan completed.")
    print(f"Output directory: {out_dir.resolve()}")
    print(f"Folders: {len(folders)}")
    print(f"Rules: {len(rule_rows)}")
    print(f"Messages scanned: {len(all_messages)}")
    print(f"Inbox unsorted: {len(unsorted)}")
    print(f"Proposed rules: {len(proposed_rules)}")


if __name__ == "__main__":
    main()
