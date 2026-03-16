# Email Deleters

A collection of tools for bulk email management across Gmail, Exchange on-premises, and Exchange Online.

---

## `deleter.gs` — Gmail (Google Apps Script)

Bulk-deletes Gmail threads matching a search query.

![image](https://github.com/user-attachments/assets/02c6dbd3-f288-4ff8-851e-6e667b3aa15b)

**Instructions:**
1. Navigate to https://script.google.com/ and create a project
2. Copy-paste `deleter.gs` into the code editor
3. Click the run button
4. Allow necessary permissions
5. Wait for the script to finish

The search query is defined in the first line — you can test it in Gmail search before running. Batch sizes and sleep timings are also configurable.

> **Warning: this script DELETES EMAIL permanently.**

![image](https://github.com/user-attachments/assets/393bbffd-be63-4c57-98ed-a1a48134a502)

---

## `exchange_ews.py` — Exchange on-premises (EWS)

Connects to an on-premises Exchange server via Exchange Web Services (EWS).

**Requirements:** `pip install exchangelib`

**Subcommands:**

```
# List all mail folders sorted by size
python exchange_ews.py --server https://mail.example.com/EWS/Exchange.asmx \
    --username DOMAIN\user --password pass --email user@example.com audit [--top N] [--skip-empty]

# Show largest messages in a folder
python exchange_ews.py ... biggest --folder "Inbox" [--top 20]

# Delete all messages from a folder (folder itself is kept)
python exchange_ews.py ... delete --folder "Deleted Items" [--yes] [--batch-size 100]
```

---

## `exchange_online.py` — Exchange Online / Microsoft 365 (Graph API)

Connects to Exchange Online via the Microsoft Graph API using MSAL device-flow authentication (no stored credentials).

**Requirements:** `pip install requests msal`

**Setup:** Register an Azure app with `Mail.ReadWrite` delegated permission and note the client ID and tenant ID.

**Subcommands:**

```
# List all mail folders sorted by size
python exchange_online.py --tenant <tenant-id> --client-id <client-id> audit [--top N]

# Show largest messages in a folder
python exchange_online.py ... biggest --folder "Inbox" [--top 20]

# Delete all messages from a folder (folder itself is kept)
python exchange_online.py ... delete --folder "Deleted Items" [--yes] [--batch-size 100]
```

Tenant ID and client ID can also be set via environment variables `M365_TENANT_ID` and `M365_CLIENT_ID`.

> **Warning: the delete subcommand permanently removes email.**
