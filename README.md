# Outlook Mail Audit

[简体中文](README.zh-CN.md) · [Security / 安全说明](SECURITY.md)

A small personal utility for inspecting Outlook folders, inbox rules and uncategorized inbox senders. It writes local CSV/JSON reports and suggests sender rules for manual review. It never creates or applies rules, moves messages, sends mail or deletes mail.

This repository archives a practical personal workflow. It is not an official Microsoft product or a complete mailbox compliance/audit solution.

## Two ways to scan

| Backend | Requirements | Authentication | Scope |
| --- | --- | --- | --- |
| Microsoft Graph | Python 3.11+, Microsoft mailbox, your own app registration | Interactive device login | The signed-in account |
| Classic Outlook | Windows, Python 3.11+, classic desktop Outlook already configured | Existing Outlook profile | All stores mounted in that profile |

The new Outlook for Windows does not provide the classic COM interface. Use Graph when classic Outlook is unavailable. Help and offline tests need only Python; dependencies load only when a real scan starts.

## Install

```powershell
git clone https://github.com/lenkazuma/outlook-mail-audit.git
cd outlook-mail-audit
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

On macOS/Linux, activate with `source .venv/bin/activate`; only Graph is supported there.

### Option A: classic Outlook on Windows

Open classic Outlook and confirm the intended profile is selected. Install the COM dependency and run:

```powershell
python -m pip install -r requirements-outlook.txt
python scan_outlook_local_readonly.py --days 365 --max-messages-per-folder 5000
```

All mounted stores are read, including any shared/archive stores your profile exposes. Reports go to `output_local/`. No Azure/Entra registration is needed for this route.

### Option B: Microsoft Graph

1. Create your own application in [Microsoft Entra app registrations](https://entra.microsoft.com/). Select supported account types matching the accounts you intend to use (include personal Microsoft accounts if needed).
2. Under **Authentication**, enable **Allow public client flows** for device login. No client secret is used by this tool.
3. Add Microsoft Graph **delegated** permissions: `User.Read`, `Mail.Read`, `MailboxSettings.Read`. Consent requirements depend on your organization's policy.
4. Copy the application/client ID into your local configuration:

```powershell
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

Edit `.env`:

```dotenv
GRAPH_CLIENT_ID=your_app_client_id
GRAPH_TENANT=common
GRAPH_SCOPES=User.Read Mail.Read MailboxSettings.Read
```

Replace the placeholder with your application's client ID. `common` supports the account types configured for your application; use your tenant ID when required by organizational policy. MSAL supplies reserved OpenID scopes automatically; do not add `offline_access`, `openid` or `profile` to the scope list.

```powershell
python scan_readonly.py --days 365 --max-messages-per-folder 5000
```

Follow the device-login message and verify the account and application shown by Microsoft. Reports go to `output/`. For separate accounts, use separate caches, for example `--token-cache .token_cache_work.bin`; the scanner refuses to silently pick from multiple cached accounts.

Official references: [MSAL device login](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens), [inbox rule read permissions](https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messagerules?view=graph-rest-1.0), [mail folder identifiers](https://learn.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0).

## Options and interpretation

```powershell
python scan_readonly.py --help
python scan_outlook_local_readonly.py --help
```

Both scanners accept `--days` (default 365), `--max-messages-per-folder` (default 5000), `--suggest-threshold` (default 5) and `--output-dir`. Integer options must be positive. Graph additionally accepts `--token-cache`.

- “Unsorted” means a message is directly in the default Inbox and has no category; it is a heuristic, not a claim that a rule failed.
- Inbox detection uses folder IDs, so a localized Inbox name is supported. Inbox subfolders are not counted as the Inbox.
- Date limits use each message's received time. Counts describe the scanned sample, limited by both time and the per-folder cap; folder totals can therefore differ.
- Suggested rules are JSON suggestions only. Graph also checks sender tokens in existing enabled rules, but this is not a full simulation of conditions, exceptions, ordering or rule execution. Local Outlook rule summaries are less detailed and do not exclude senders covered by existing rules.

## Local reports

| File | Contents |
| --- | --- |
| `mailbox_profile.json` | Scan time, account/profile and counts; local backend also records some skipped-store/folder errors |
| `folders.csv` | Folder names and counts |
| `rules_inventory.csv` | Existing rule summaries |
| `all_messages_metadata.csv` | Subject, sender, dates, category and identifiers of sampled messages |
| `unsorted_inbox_messages.csv` | Uncategorized messages directly in the Inbox |
| `unsorted_senders.csv` | Sender counts for that subset |
| `proposed_rules.json` | Suggestions to review manually |
| `folders.json`, `rules.json` | Graph-only raw folder/rule records |

Mail bodies and attachment contents are not exported. Rule definitions may contain sensitive condition text. Empty CSV reports are empty files. CSV text beginning with spreadsheet formula characters is prefixed with an apostrophe for safer viewing; JSON values are unchanged. The backend CSV schemas differ because the APIs expose different metadata.

[examples/synthetic](examples/synthetic) contains deliberately fictional reports made without connecting to a mailbox. No real mailbox data belongs in this repository.

## Privacy and read-only boundary

Mailbox API requests are GET-only. Graph scopes are restricted to the three read permissions above; requests and pagination are restricted to `https://graph.microsoft.com/v1.0/` and redirects are disabled. Sign-in itself communicates with Microsoft and writes a local token cache. Local scans read Outlook's existing profile and write report files; Outlook may do its own normal background synchronization.

`.env`, token caches, standard output folders and report filenames are ignored by Git. **Custom output locations still need your review before any commit.** Keep reports and token caches on a private local disk, outside public/shared or automatically synced folders. The cache is plaintext, not encrypted; the script's POSIX file mode is not a substitute for Windows folder ACLs. Removing a cache logs this tool out locally but does not revoke already issued tokens.

## Tests

```powershell
python -m unittest discover -s tests -v
```

Tests use synthetic data and mocks only: no Outlook application, Microsoft login, network or mail changes. CI checks the same offline tests on Windows/Linux and Python 3.11/3.13.

## Status and remaining work

Prepared for public archiving: bilingual documentation, MIT license, split dependencies, synthetic examples, CSV escaping, localized inbox detection, Graph host restrictions and offline tests.

Still open:

- [ ] Verify a fresh Graph sign-in and classic Outlook scan in representative environments (not exercised by offline CI).
- [ ] Add Graph throttling/backoff and long-running access-token refresh.
- [ ] Add explicit store selection and better Exchange address-to-SMTP resolution for local Outlook.
- [ ] Expand rule coverage/exception handling and clearly report every incomplete local rule read.
- [ ] Add structured progress/error reporting and dependency update checks.

The tool currently stops on unhandled Graph errors, including throttling or token expiry; do not interpret a failed run or previously existing report as a new complete scan. Local Outlook can expose internal Exchange sender addresses, and some protected folders/rule properties may be inaccessible. Suggestions always need manual review.

## License

[MIT](LICENSE), copyright 2026 Zhenzhen Liu.
