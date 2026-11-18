"""Check the Microsoft Graph connection before using the tool.   python graph_test.py
Uses config.json (or BC_CONFIG) for graph_tenant / graph_client_id / graph_mailbox, and the environment variable
BC_GRAPH_SECRET for the secret. It only READS: signs in, then lists the 5 newest mails (subject + sender)."""
import json
import os
from pathlib import Path

from bc import graph

cfg = json.loads(Path(os.environ.get("BC_CONFIG", Path(__file__).parent / "config.json")).read_text())
secret = os.environ.get("BC_GRAPH_SECRET") or cfg.get("graph_secret", "")
if not secret:
    raise SystemExit("Set the secret first.  Windows:  set BC_GRAPH_SECRET=your-secret   Mac/Linux:  export BC_GRAPH_SECRET=your-secret")
g = graph.Graph(cfg["graph_tenant"], cfg["graph_client_id"], secret, cfg["graph_mailbox"])
page = g.call("GET", f"/users/{cfg['graph_mailbox']}/mailFolders/inbox/messages?$top=5&$select=subject,from,receivedDateTime")
print("Signed in OK. Newest mails:")
for m in page.get("value", []):
    print(" -", m.get("receivedDateTime", "")[:16], "|", (m.get("from") or {}).get("emailAddress", {}).get("address", ""), "|", m.get("subject"))
