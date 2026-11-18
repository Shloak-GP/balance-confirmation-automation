# Microsoft Graph setup, step by step

## A. Register the app (needs an Entra admin; if you are not one, send IT_REQUEST.md)
1. Go to https://entra.microsoft.com, App registrations, New registration. Name: "Balance check". Single tenant. Register.
2. On the Overview page copy **Directory (tenant) ID** and **Application (client) ID**.
3. Certificates & secrets, New client secret. Copy the secret VALUE now (it is shown once).
4. API permissions, Add a permission, Microsoft Graph, **Application permissions**: tick `Mail.Read`, `Mail.ReadWrite` and `Files.Read.All` (the last is for the OneDrive folder). Then "Grant admin consent".
5. Ask the Exchange admin to limit the app to the one mailbox (application access policy), so it cannot read others.

## B. Test the connection
1. In `config.json` set:  "graph_tenant": "<tenant id>", "graph_client_id": "<client id>", "graph_mailbox": "owner@yourcompany.com", "thread_subject": "<exact subject of the daily MIS thread>"
2. Put the secret in the environment (never in the file):  Windows  `set BC_GRAPH_SECRET=...`   Mac/Linux  `export BC_GRAPH_SECRET=...`
3. Run `python graph_test.py`. You should see "Signed in OK" and the 5 newest mails. A 403 means the permission or consent is missing.

## C. Run the tool
`python web.py` (or `python run_real.py` for the real-data pack, with the keys in local_overrides.json). On Home press **Pull from mailbox**.
On the MIS page press **Reply on the thread (Outlook draft, file attached)**. The draft opens; he reviews and sends.

On a server: set BC_GRAPH_SECRET in the service's environment (or the container's `BC_GRAPH_SECRET`).

## D. Reading the shared OneDrive folder by its link
In OneDrive/SharePoint, right-click the folder, Share, copy the link (people in the company, view only is enough).
In `config.json` set `"onedrive_link": "<that link>"`. The tool copies NEW and CHANGED files into its own mirror folder (read only:
nothing in OneDrive is changed or deleted) when Home opens, every 10 minutes if `auto_scan_minutes` is set, or when you press
**Sync OneDrive folder**. The mirror is read like any other reply folder.
