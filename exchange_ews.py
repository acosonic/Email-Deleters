#!/usr/bin/env python3
import argparse
import sys

from exchangelib import Account, Configuration, Credentials, DELEGATE, FaultTolerance
from exchangelib.folders import Folder


MAIL_CONTAINER_CLASSES = {
    "IPF.Note",
    "IPF.Note.OutlookHomepage",
}


def human_size(num):
    units = ["B", "KB", "MB", "GB", "TB"]
    n = float(num or 0)
    for unit in units:
        if n < 1024 or unit == units[-1]:
            return f"{n:.1f} {unit}"
        n /= 1024.0


def connect(server, username, password, email):
    creds = Credentials(username=username, password=password)

    config = Configuration(
        service_endpoint=server,
        credentials=creds,
        retry_policy=FaultTolerance(max_wait=60),
    )

    return Account(
        primary_smtp_address=email,
        config=config,
        autodiscover=False,
        access_type=DELEGATE,
    )


def iter_folders(root_folder):
    yield root_folder
    for child in getattr(root_folder, "children", []):
        yield from iter_folders(child)


def folder_path(folder):
    parts = []
    cur = folder
    while cur is not None:
        name = getattr(cur, "name", None)
        if name:
            parts.append(name)
        cur = getattr(cur, "parent", None)
    parts.reverse()
    return "/".join(parts)


def is_mail_folder(folder):
    folder_class = getattr(folder, "folder_class", None)
    return folder_class in MAIL_CONTAINER_CLASSES


def scan_folder(folder, top_n=10, include_biggest=True):
    total_size = 0
    total_count = 0
    biggest = []

    try:
        if include_biggest:
            qs = folder.all().only(
                "subject",
                "size",
                "datetime_received",
                "sender",
                "has_attachments",
            )
        else:
            qs = folder.all().only("size")

        for item in qs:
            size = int(getattr(item, "size", 0) or 0)
            total_size += size
            total_count += 1

            if include_biggest:
                biggest.append({
                    "size": size,
                    "subject": getattr(item, "subject", "") or "",
                    "received": getattr(item, "datetime_received", None),
                    "sender": str(getattr(item, "sender", "") or ""),
                    "has_attachments": getattr(item, "has_attachments", False),
                })

    except Exception as e:
        return {
            "error": str(e),
            "count": 0,
            "size": 0,
            "biggest": [],
        }

    if include_biggest:
        biggest.sort(key=lambda x: x["size"], reverse=True)
        biggest = biggest[:top_n]
    else:
        biggest = []

    return {
        "error": None,
        "count": total_count,
        "size": total_size,
        "biggest": biggest,
    }


def cmd_audit(account, limit, skip_empty):
    rows = []

    for folder in iter_folders(account.root):
        if not isinstance(folder, Folder):
            continue
        if not is_mail_folder(folder):
            continue

        info = scan_folder(folder, include_biggest=False)
        if info["error"]:
            print(f"[WARN] {folder_path(folder)} -> {info['error']}", file=sys.stderr)
            continue

        if skip_empty and info["count"] == 0:
            continue

        rows.append({
            "path": folder_path(folder),
            "count": info["count"],
            "size": info["size"],
        })

    rows.sort(key=lambda x: x["size"], reverse=True)

    print(f"{'Size':>12}  {'Count':>8}  Folder")
    print("-" * 100)
    for row in rows[:limit] if limit else rows:
        print(f"{human_size(row['size']):>12}  {row['count']:>8}  {row['path']}")


def find_folder_by_name(account, name):
    wanted = name.lower()
    for folder in iter_folders(account.root):
        if not isinstance(folder, Folder):
            continue
        if not is_mail_folder(folder):
            continue

        p = folder_path(folder).lower()
        n = (getattr(folder, "name", "") or "").lower()

        if p == wanted or n == wanted or p.endswith("/" + wanted):
            return folder
    return None


def cmd_biggest(account, folder_name, top_n):
    folder = find_folder_by_name(account, folder_name)
    if not folder:
        print(f"Folder not found: {folder_name}", file=sys.stderr)
        sys.exit(2)

    info = scan_folder(folder, top_n=top_n, include_biggest=True)
    if info["error"]:
        print(f"Error in folder {folder_path(folder)}: {info['error']}", file=sys.stderr)
        sys.exit(3)

    print(f"Folder: {folder_path(folder)}")
    print(f"Total: {info['count']} messages, {human_size(info['size'])}")
    print("-" * 140)
    print(f"{'Size':>10}  {'Att':>3}  {'Received':<25}  {'Sender':<35}  Subject")
    print("-" * 140)

    for m in info["biggest"]:
        recv = str(m["received"] or "-")
        att = "yes" if m["has_attachments"] else "no"
        sender = (m["sender"] or "-")[:35]
        subject = (m["subject"] or "").replace("\n", " ")[:60]
        print(f"{human_size(m['size']):>10}  {att:>3}  {recv:<25}  {sender:<35}  {subject}")


def confirm(prompt):
    ans = input(f"{prompt} [yes/no]: ").strip().lower()
    return ans == "yes"


def cmd_delete(account, folder_name, assume_yes=False, batch_size=100):
    folder = find_folder_by_name(account, folder_name)
    if not folder:
        print(f"Folder not found: {folder_name}", file=sys.stderr)
        sys.exit(2)

    info = scan_folder(folder, include_biggest=False)
    if info["error"]:
        print(f"Error in folder {folder_path(folder)}: {info['error']}", file=sys.stderr)
        sys.exit(3)

    print(f"Target folder: {folder_path(folder)}")
    print(f"Messages to delete: {info['count']}")
    print(f"Total size: {human_size(info['size'])}")

    if info["count"] == 0:
        print("Folder is empty.")
        return

    if not assume_yes:
        ok = confirm("Are you sure you want to delete ALL messages from this folder? The folder itself will be kept.")
        if not ok:
            print("Cancelled.")
            return

    deleted = 0
    errors = 0

    # If already in Deleted Items, hard-delete. Otherwise move to Deleted Items.
    target_is_trash = folder == account.trash or folder.name.lower() == "deleted items"

    qs = folder.all().only("subject")

    batch = []
    for item in qs:
        batch.append(item)
        if len(batch) >= batch_size:
            d, e = process_delete_batch(account, batch, target_is_trash)
            deleted += d
            errors += e
            print(f"Processed: {deleted + errors}, succeeded: {deleted}, errors: {errors}")
            batch = []

    if batch:
        d, e = process_delete_batch(account, batch, target_is_trash)
        deleted += d
        errors += e

    print("-" * 60)
    print(f"Done. Succeeded: {deleted}, errors: {errors}")


def process_delete_batch(account, items, target_is_trash):
    deleted = 0
    errors = 0

    for item in items:
        try:
            if target_is_trash:
                item.delete()
            else:
                item.move(to_folder=account.trash)
            deleted += 1
        except Exception as e:
            errors += 1
            subj = getattr(item, "subject", "") or ""
            print(f"[WARN] Failed to delete/move: {subj[:60]} -> {e}", file=sys.stderr)

    return deleted, errors


def main():
    p = argparse.ArgumentParser(description="EWS mailbox size inspector")
    p.add_argument("--server", required=True, help="EWS URL")
    p.add_argument("--username", required=True, help="DOMAIN\\user or email")
    p.add_argument("--password", required=True, help="Password")
    p.add_argument("--email", required=True, help="Primary SMTP address of the mailbox")

    sub = p.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("audit", help="List mail folders sorted by size")
    a.add_argument("--top", type=int, default=None, help="Show only top N folders")
    a.add_argument("--skip-empty", action="store_true", help="Skip empty folders")

    b = sub.add_parser("biggest", help="Show largest messages in a folder")
    b.add_argument("--folder", required=True, help="Folder name or full path")
    b.add_argument("--top", type=int, default=20, help="Number of messages to show")

    d = sub.add_parser("delete", help="Delete all messages from a folder, keep the folder itself")
    d.add_argument("--folder", required=True, help='Folder name, e.g. "Deleted Items"')
    d.add_argument("--yes", action="store_true", help="Skip confirmation prompt")
    d.add_argument("--batch-size", type=int, default=100, help="Messages per batch")

    args = p.parse_args()

    account = connect(
        server=args.server,
        username=args.username,
        password=args.password,
        email=args.email,
    )

    if args.cmd == "audit":
        cmd_audit(account, args.top, args.skip_empty)
    elif args.cmd == "biggest":
        cmd_biggest(account, args.folder, args.top)
    elif args.cmd == "delete":
        cmd_delete(account, args.folder, assume_yes=args.yes, batch_size=args.batch_size)


if __name__ == "__main__":
    main()
