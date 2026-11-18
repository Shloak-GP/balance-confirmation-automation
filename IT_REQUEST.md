# Request to IT: Balance-confirmation tool (internal finance tool)

What it does: reads customer balance-confirmation replies from ONE finance mailbox, lets the finance owner check
them in a web page, and creates a draft reply (never sent) on the daily MIS thread.

Please provide:
1. **A Microsoft Entra app registration** ("Balance check") with a client secret, and the Tenant ID + Client ID.
   Application permissions (admin consent): `Mail.Read`, `Mail.ReadWrite` and `Files.Read.All` (or `Sites.Selected`).
   Please LIMIT the app to the one mailbox with an Exchange application access policy
   (New-ApplicationAccessPolicy -AccessRight RestrictAccess ...), so it cannot read any other mailbox.
   - Mail.Read: read replies from that mailbox (read only; nothing is moved, marked read or deleted).
   - Files.Read.All (or, better, Sites.Selected limited to the one site/library): read the finance reply folder on OneDrive/SharePoint by its sharing link. Read only.
   - Mail.ReadWrite: create the DRAFT reply on the MIS thread with the MIS file attached. The tool never sends mail.
2. **A server (or container host)** to run the tool, with port 8080 open to the finance team only, and outbound access to
   login.microsoftonline.com and graph.microsoft.com.
3. **The sharing link** of the OneDrive/SharePoint folder where the finance team saves replies, plus a local or shared folder for the tool's own data and the MIS reports.

Secrets are passed as environment variables (BC_GRAPH_SECRET), not stored in the config file.
