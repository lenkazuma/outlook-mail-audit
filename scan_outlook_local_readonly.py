from __future__ import annotations

import argparse
import json
from collections import Counter
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

from audit_utils import positive_int, write_csv


def safe_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def collect_folders(root_folder, parent_path: str = "") -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    folder_path = f"{parent_path}/{safe_str(root_folder.Name)}" if parent_path else safe_str(root_folder.Name)
    row = {
        "entryId": safe_str(getattr(root_folder, "EntryID", "")),
        "storeId": safe_str(getattr(root_folder, "StoreID", "")),
        "name": safe_str(root_folder.Name),
        "path": folder_path,
        "itemCount": int(getattr(root_folder, "Items", []).Count) if getattr(root_folder, "Items", None) else 0,
        "unreadItemCount": int(getattr(root_folder, "UnReadItemCount", 0)),
    }
    rows.append(row)
    for i in range(1, root_folder.Folders.Count + 1):
        child = root_folder.Folders.Item(i)
        rows.extend(collect_folders(child, folder_path))
    return rows


def collect_rules(inbox_folder) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    try:
        rules = inbox_folder.Store.GetRules()
    except Exception:
        return rows

    for i in range(1, rules.Count + 1):
        rule = rules.Item(i)
        row = {
            "name": safe_str(getattr(rule, "Name", "")),
            "enabled": bool(getattr(rule, "Enabled", False)),
            "executionOrder": int(getattr(rule, "ExecutionOrder", 0)),
            "isLocalRule": bool(getattr(rule, "IsLocalRule", False)),
            "conditionsSummary": "",
            "actionsSummary": "",
        }

        conds = []
        actions = []
        try:
            if rule.Conditions.From.Enabled:
                conds.append("From")
            if rule.Conditions.Subject.Enabled:
                conds.append("Subject")
            if rule.Conditions.Body.Enabled:
                conds.append("Body")
            if rule.Conditions.SentTo.Enabled:
                conds.append("SentTo")
            if rule.Conditions.Categories.Enabled:
                conds.append("Categories")
        except Exception:
            pass

        try:
            if rule.Actions.MoveToFolder.Enabled:
                folder_name = safe_str(getattr(rule.Actions.MoveToFolder.Folder, "Name", ""))
                actions.append(f"MoveToFolder:{folder_name}")
            if rule.Actions.AssignToCategory.Enabled:
                actions.append("AssignToCategory")
            if rule.Actions.Stop.Enabled:
                actions.append("Stop")
            if rule.Actions.Delete.Enabled:
                actions.append("Delete")
        except Exception:
            pass

        row["conditionsSummary"] = "; ".join(conds)
        row["actionsSummary"] = "; ".join(actions)
        rows.append(row)
    return rows


def collect_messages_from_folder(folder, folder_path: str, cutoff: datetime, max_messages: int) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    items = folder.Items
    sorted_by_received = False
    try:
        items.Sort("[ReceivedTime]", True)
        sorted_by_received = True
    except Exception:
        pass

    count = items.Count
    scanned = 0
    for i in range(1, count + 1):
        if scanned >= max_messages:
            break
        item = items.Item(i)
        if getattr(item, "Class", 0) != 43:
            continue
        received = getattr(item, "ReceivedTime", None)
        if not received:
            continue
        if not isinstance(received, datetime):
            continue
        # COM may return aware datetimes; naive datetimes are local wall times.
        if received.astimezone(UTC) < cutoff.astimezone(UTC):
            if sorted_by_received:
                break
            continue

        sender_email = safe_str(getattr(item, "SenderEmailAddress", ""))
        sender_name = safe_str(getattr(item, "SenderName", ""))
        categories = safe_str(getattr(item, "Categories", ""))
        rows.append(
            {
                "entryId": safe_str(getattr(item, "EntryID", "")),
                "folderPath": folder_path,
                "folderEntryId": safe_str(getattr(folder, "EntryID", "")),
                "storeId": safe_str(getattr(folder, "StoreID", "")),
                "subject": safe_str(getattr(item, "Subject", "")).replace("\n", " "),
                "receivedTime": received.isoformat() if isinstance(received, datetime) else safe_str(received),
                "senderName": sender_name,
                "senderEmail": sender_email.lower(),
                "categories": categories,
                "isRead": bool(getattr(item, "UnRead", False)) is False,
                "importance": int(getattr(item, "Importance", 1)),
                "hasAttachments": bool(getattr(item, "Attachments", []).Count > 0)
                if getattr(item, "Attachments", None)
                else False,
            }
        )
        scanned += 1
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description="Read-only scan via local Outlook COM")
    parser.add_argument("--days", type=positive_int, default=365, help="Scan recent N days")
    parser.add_argument("--max-messages-per-folder", type=positive_int, default=5000, help="Safety cap per folder")
    parser.add_argument("--suggest-threshold", type=positive_int, default=5, help="Min count for sender suggestions")
    parser.add_argument("--output-dir", default="output_local", help="Output directory")
    args = parser.parse_args()

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Classic Outlook on Windows required: python -m pip install -r requirements-outlook.txt") from exc

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    cutoff = datetime.now(UTC) - timedelta(days=args.days)

    pythoncom.CoInitialize()
    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    namespace = getattr(outlook, "Session", None)
    if namespace is None:
        try:
            namespace = outlook.GetNamespace("MAPI")
        except Exception as exc:
            raise RuntimeError(
                "无法访问 Outlook MAPI。请确认安装并打开的是经典版 Outlook（非新版 Outlook）。"
            ) from exc

    all_folders: List[Dict[str, Any]] = []
    all_messages: List[Dict[str, Any]] = []
    all_rules: List[Dict[str, Any]] = []
    profile_name = ""
    inbox_ids = set()
    scan_errors = []

    for store_index in range(1, namespace.Stores.Count + 1):
        store = namespace.Stores.Item(store_index)
        profile_name = profile_name or safe_str(getattr(store, "DisplayName", ""))
        root = store.GetRootFolder()
        folders = collect_folders(root)
        all_folders.extend(folders)

        try:
            inbox = store.GetDefaultFolder(6)
            inbox_ids.add((safe_str(inbox.EntryID), safe_str(inbox.StoreID)))
            rules = collect_rules(inbox)
            for r in rules:
                r["storeName"] = safe_str(getattr(store, "DisplayName", ""))
            all_rules.extend(rules)
        except Exception as exc:
            scan_errors.append({"store": safe_str(store.DisplayName), "stage": "inbox/rules", "error": type(exc).__name__})

        for fr in folders:
            try:
                folder_obj = namespace.GetFolderFromID(fr["entryId"], fr["storeId"])
            except Exception as exc:
                scan_errors.append({"folder": fr["path"], "stage": "open", "error": type(exc).__name__})
                continue
            rows = collect_messages_from_folder(
                folder=folder_obj,
                folder_path=fr["path"],
                cutoff=cutoff,
                max_messages=args.max_messages_per_folder,
            )
            all_messages.extend(rows)

    inbox_messages = [m for m in all_messages if (m["folderEntryId"], m["storeId"]) in inbox_ids]
    unsorted = [m for m in inbox_messages if not m["categories"]]
    sender_counter = Counter([m["senderEmail"] for m in unsorted if m["senderEmail"]])

    sender_rows = [
        {"sender": sender, "unsortedCount": count}
        for sender, count in sender_counter.most_common()
    ]
    proposed_rules = []
    for row in sender_rows:
        if row["unsortedCount"] < args.suggest_threshold:
            continue
        sender = row["sender"]
        domain = sender.split("@")[-1] if "@" in sender else "unknown-domain"
        proposed_rules.append(
            {
                "type": "sender",
                "match": sender,
                "reason": f"Inbox unsorted {row['unsortedCount']} messages in last {args.days} days",
                "suggestedFolder": f"Auto/{domain}",
            }
        )

    (out_dir / "mailbox_profile.json").write_text(
        json.dumps(
            {
                "scannedAtLocal": datetime.now().isoformat(),
                "profileDisplayName": profile_name,
                "days": args.days,
                "foldersScanned": len(all_folders),
                "messagesScanned": len(all_messages),
                "inboxUnsortedCount": len(unsorted),
                "rulesCount": len(all_rules),
                "proposedRuleCount": len(proposed_rules),
                "scanErrors": scan_errors,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    (out_dir / "proposed_rules.json").write_text(json.dumps(proposed_rules, ensure_ascii=False, indent=2), encoding="utf-8")

    write_csv(out_dir / "folders.csv", all_folders)
    write_csv(out_dir / "rules_inventory.csv", all_rules)
    write_csv(out_dir / "all_messages_metadata.csv", all_messages)
    write_csv(out_dir / "unsorted_inbox_messages.csv", unsorted)
    write_csv(out_dir / "unsorted_senders.csv", sender_rows)

    pythoncom.CoUninitialize()
    print("Local Outlook read-only scan completed.")
    print(f"Skipped folders/store errors: {len(scan_errors)} (see mailbox_profile.json)")
    print(f"Output directory: {out_dir.resolve()}")
    print(f"Folders: {len(all_folders)}")
    print(f"Rules: {len(all_rules)}")
    print(f"Messages scanned: {len(all_messages)}")
    print(f"Inbox unsorted: {len(unsorted)}")
    print(f"Proposed rules: {len(proposed_rules)}")


if __name__ == "__main__":
    main()
