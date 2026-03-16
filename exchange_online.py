#!/usr/bin/env python3
import os
import sys
import argparse
import requests
import msal
from collections import deque

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.ReadWrite"]

def human_size(num):
    units = ["B", "KB", "MB", "GB", "TB"]
    n = float(num)
    for unit in units:
        if n < 1024 or unit == units[-1]:
            return f"{n:.1f} {unit}"
        n /= 1024.0

class GraphClient:
    def __init__(self, tenant_id, client_id):
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.app = msal.PublicClientApplication(client_id=client_id, authority=authority)
        self.token = None

    def login(self):
        accounts = self.app.get_accounts()
        result = None
        if accounts:
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])

        if not result:
            flow = self.app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(f"Device flow failed: {flow}")
            print(flow["message"])
            result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError(f"Login failed: {result}")

        self.token = result["access_token"]

    def get(self, path, params=None):
        headers = {"Authorization": f"Bearer {self.token}"}
        url = f"{GRAPH_BASE}{path}"
        r = requests.get(url, headers=headers, params=params, timeout=60)
        if r.status_code >= 400:
            raise RuntimeError(f"GET {url} failed: {r.status_code} {r.text}")
        return r.json()

    def post(self, path, json=None):
        headers = {"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"}
        url = f"{GRAPH_BASE}{path}"
        r = requests.post(url, headers=headers, json=json, timeout=60)
        if r.status_code >= 400:
            raise RuntimeError(f"POST {url} failed: {r.status_code} {r.text}")
        return r.json() if r.content else {}

    def delete(self, path):
        headers = {"Authorization": f"Bearer {self.token}"}
        url = f"{GRAPH_BASE}{path}"
        r = requests.delete(url, headers=headers, timeout=60)
        if r.status_code >= 400:
            raise RuntimeError(f"DELETE {url} failed: {r.status_code} {r.text}")

    def get_all(self, path, params=None):
        headers = {"Authorization": f"Bearer {self.token}"}
        url = f"{GRAPH_BASE}{path}"
        out = []

        while url:
            r = requests.get(url, headers=headers, params=params if url.startswith(GRAPH_BASE) else None, timeout=60)
            if r.status_code >= 400:
                raise RuntimeError(f"GET {url} failed: {r.status_code} {r.text}")
            data = r.json()
            out.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            params = None

        return out

def list_all_folders(gc):
    root_folders = gc.get_all(
        "/me/mailFolders",
        params={
            "$top": 100,
            "$select": "id,displayName,childFolderCount,totalItemCount"
        }
    )

    all_folders = []
    queue = deque()

    for f in root_folders:
        f["_path"] = f["displayName"]
        all_folders.append(f)
        queue.append(f)

    while queue:
        parent = queue.popleft()
        if parent.get("childFolderCount", 0) > 0:
            children = gc.get_all(
                f"/me/mailFolders/{parent['id']}/childFolders",
                params={
                    "$top": 100,
                    "$select": "id,displayName,childFolderCount,totalItemCount"
                }
            )
            for c in children:
                c["_path"] = f"{parent['_path']}/{c['displayName']}"
                all_folders.append(c)
                queue.append(c)

    return all_folders

def estimate_folder_size(gc, folder_id):
    total_size = 0
    total_count = 0

    messages = gc.get_all(
        f"/me/mailFolders/{folder_id}/messages",
        params={
            "$top": 200,
            "$select": "id,size"
        }
    )

    for m in messages:
        total_size += int(m.get("size", 0))
        total_count += 1

    return total_size, total_count

def audit_folders(gc, limit=None):
    folders = list_all_folders(gc)
    rows = []

    for f in folders:
        try:
            size_bytes, msg_count = estimate_folder_size(gc, f["id"])
        except Exception as e:
            print(f"[WARN] {f['_path']}: {e}", file=sys.stderr)
            size_bytes, msg_count = 0, 0

        rows.append({
            "path": f["_path"],
            "items_reported": f.get("totalItemCount", 0),
            "items_scanned": msg_count,
            "size_bytes": size_bytes,
        })

    rows.sort(key=lambda x: x["size_bytes"], reverse=True)

    print(f"{'Size':>12}  {'Scanned':>8}  {'Reported':>8}  Folder")
    print("-" * 90)
    for row in rows[:limit] if limit else rows:
        print(f"{human_size(row['size_bytes']):>12}  {row['items_scanned']:>8}  {row['items_reported']:>8}  {row['path']}")

def find_folder_by_name(gc, folder_name):
    folders = list_all_folders(gc)
    folder_name_lower = folder_name.lower()
    for f in folders:
        if f["_path"].lower() == folder_name_lower or f["displayName"].lower() == folder_name_lower:
            return f
    return None


def confirm(prompt):
    ans = input(f"{prompt} [yes/no]: ").strip().lower()
    return ans == "yes"


def cmd_delete(gc, folder_name, assume_yes=False, batch_size=100):
    target = find_folder_by_name(gc, folder_name)
    if not target:
        print(f"Folder not found: {folder_name}", file=sys.stderr)
        sys.exit(2)

    size_bytes, msg_count = estimate_folder_size(gc, target["id"])

    print(f"Target folder: {target['_path']}")
    print(f"Messages to delete: {msg_count}")
    print(f"Total size: {human_size(size_bytes)}")

    if msg_count == 0:
        print("Folder is empty.")
        return

    if not assume_yes:
        ok = confirm("Are you sure you want to delete ALL messages from this folder? The folder itself will be kept.")
        if not ok:
            print("Cancelled.")
            return

    target_is_trash = target["displayName"].lower() in ("deleted items", "trash")

    messages = gc.get_all(
        f"/me/mailFolders/{target['id']}/messages",
        params={"$top": 200, "$select": "id,subject"}
    )

    deleted = 0
    errors = 0

    for i in range(0, len(messages), batch_size):
        batch = messages[i:i + batch_size]
        for msg in batch:
            try:
                if target_is_trash:
                    gc.delete(f"/me/messages/{msg['id']}")
                else:
                    gc.post(f"/me/messages/{msg['id']}/move", json={"destinationId": "deleteditems"})
                deleted += 1
            except Exception as e:
                errors += 1
                subj = (msg.get("subject") or "")[:60]
                print(f"[WARN] Failed to delete/move: {subj} -> {e}", file=sys.stderr)
        print(f"Processed: {deleted + errors}, succeeded: {deleted}, errors: {errors}")

    print("-" * 60)
    print(f"Done. Succeeded: {deleted}, errors: {errors}")


def show_biggest_messages(gc, folder_name, top_n=20):
    target = find_folder_by_name(gc, folder_name)

    if not target:
        raise RuntimeError(f"Folder not found: {folder_name}")

    messages = gc.get_all(
        f"/me/mailFolders/{target['id']}/messages",
        params={"$top": 200, "$select": "subject,receivedDateTime,from,size,hasAttachments"}
    )

    messages.sort(key=lambda x: int(x.get("size", 0)), reverse=True)

    print(f"Top {top_n} largest messages in: {target['_path']}")
    print("-" * 120)
    for m in messages[:top_n]:
        sender = ""
        frm = m.get("from") or {}
        email = frm.get("emailAddress") or {}
        sender = email.get("address", "-")
        subject = (m.get("subject") or "").replace("\n", " ").strip()
        received = m.get("receivedDateTime", "-")
        attach = "yes" if m.get("hasAttachments") else "no"
        print(f"{human_size(int(m.get('size', 0))):>10}  {attach:>3}  {received}  {sender:<35}  {subject[:50]}")

def main():
    parser = argparse.ArgumentParser(description="Small Exchange Online mailbox audit tool")
    parser.add_argument("--tenant", default=os.getenv("M365_TENANT_ID", "common"),
                        help="Tenant ID or 'common'")
    parser.add_argument("--client-id", default=os.getenv("M365_CLIENT_ID"),
                        help="Azure app client ID")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_audit = sub.add_parser("audit", help="Show folders sorted by estimated size")
    p_audit.add_argument("--top", type=int, default=None, help="Show only top N folders")

    p_big = sub.add_parser("biggest", help="Show biggest messages in a folder")
    p_big.add_argument("--folder", required=True, help="Folder path or display name")
    p_big.add_argument("--top", type=int, default=20, help="How many messages")

    p_del = sub.add_parser("delete", help="Delete all messages from a folder, keep the folder itself")
    p_del.add_argument("--folder", required=True, help='Folder name or path, e.g. "Deleted Items"')
    p_del.add_argument("--yes", action="store_true", help="Skip confirmation prompt")
    p_del.add_argument("--batch-size", type=int, default=100, help="Messages per batch")

    args = parser.parse_args()

    if not args.client_id:
        print("Missing --client-id or M365_CLIENT_ID env var", file=sys.stderr)
        sys.exit(1)

    gc = GraphClient(args.tenant, args.client_id)
    gc.login()

    if args.cmd == "audit":
        audit_folders(gc, args.top)
    elif args.cmd == "biggest":
        show_biggest_messages(gc, args.folder, args.top)
    elif args.cmd == "delete":
        cmd_delete(gc, args.folder, assume_yes=args.yes, batch_size=args.batch_size)

if __name__ == "__main__":
    main()
