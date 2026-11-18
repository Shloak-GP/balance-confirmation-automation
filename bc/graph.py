"""Microsoft Graph (Microsoft 365): read replies from a mailbox and create the daily reply on the real thread.
Runs on the server. No Outlook program needed. Needs an app registration by IT with these APPLICATION permissions:
  Mail.Read       read the mailbox (pull replies)
  Mail.ReadWrite  create the draft reply on the thread (the draft is only created, never sent)
IT should limit the app to the one mailbox (Exchange 'application access policy').
The client secret is read from the environment (BC_GRAPH_SECRET), never from a file.
Mail is only READ; nothing is moved, marked read or deleted."""
import base64
import json
import re
import urllib.parse
import urllib.request
from pathlib import Path

from .outlook import _maybe, _safe, _why

ROOT = "https://graph.microsoft.com/v1.0"


def _http(method, url, headers=None, data=None):
    req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()


class _Item:
    """A Graph message that looks like an Outlook item, so the same relevance rules apply."""

    def __init__(self, m):
        self.Subject = m.get("subject") or ""
        fr = (m.get("from") or {}).get("emailAddress") or {}
        self.SenderName, self.SenderEmailAddress = fr.get("name", ""), fr.get("address", "")
        self.Body = m.get("bodyPreview") or ""
        self.Attachments = [type("A", (), {"FileName": a.get("name", "")})() for a in (m.get("attachments") or [])]
        self.Class = 43


class Graph:
    def __init__(self, tenant, client_id, secret, mailbox, req=None):
        self.tenant, self.client_id, self.secret, self.mailbox = tenant, client_id, secret, mailbox
        self.req = req or _http
        self._tok = None

    def _token(self):
        if not self._tok:
            body = urllib.parse.urlencode(dict(client_id=self.client_id, client_secret=self.secret,
                                                scope="https://graph.microsoft.com/.default",
                                                grant_type="client_credentials")).encode()
            st, raw = self.req("POST", f"https://login.microsoftonline.com/{self.tenant}/oauth2/v2.0/token",
                               {"Content-Type": "application/x-www-form-urlencoded"}, body)
            if st != 200:
                raise RuntimeError(f"Microsoft sign-in failed ({st}): {raw[:200]!r}")
            self._tok = json.loads(raw)["access_token"]
        return self._tok

    def call(self, method, path, body=None, raw=False):
        url = path if path.startswith("http") else ROOT + path
        h = {"Authorization": "Bearer " + self._token()}
        data = None
        if body is not None:
            h["Content-Type"] = "application/json"
            data = json.dumps(body).encode()
        st, out = self.req(method, url, h, data)
        if st >= 300:
            raise RuntimeError(f"Graph {method} {path.split('?')[0]} failed ({st}): {out[:300]!r}")
        return out if raw else (json.loads(out) if out else {})

    def _folder(self, folder):
        return "inbox" if not folder or folder.lower() == "inbox" else folder      # a folder id or well-known name

    def pull(self, dest, seen_ids, keywords=(), known=None, maybe_dir=None, folder="inbox", since_days=30, limit=2000):
        """Returns (saved_paths, seen_ids, report). Each mail is saved as .eml (full message with attachments)."""
        import datetime as dt
        dest = Path(dest); dest.mkdir(parents=True, exist_ok=True)
        seen = set(seen_ids)
        kws = [k.strip().lower() for k in (keywords or []) if k and k.strip()]
        rep = dict(in_folder=0, saved=0, already=0, not_mail=0, filtered=0, maybe=0, why={}, failed=[])
        saved = []
        since = (dt.datetime.now(dt.timezone.utc) - dt.timedelta(days=since_days)).strftime("%Y-%m-%dT%H:%M:%SZ")
        q = urllib.parse.urlencode({"$filter": f"receivedDateTime ge {since}", "$top": "50",
                                    "$select": "id,subject,from,hasAttachments,receivedDateTime,bodyPreview,internetMessageId",
                                    "$expand": "attachments($select=name)", "$orderby": "receivedDateTime desc"})
        url = f"/users/{urllib.parse.quote(self.mailbox)}/mailFolders/{self._folder(folder)}/messages?{q}"
        n = 0
        while url and n < limit:
            page = self.call("GET", url)
            for m in page.get("value", []):
                n += 1
                rep["in_folder"] += 1
                key = "graph|" + (m.get("internetMessageId") or m["id"])
                if key in seen:
                    rep["already"] += 1
                    continue
                try:
                    it = _Item(m)
                    why = "all"
                    if kws or known:
                        why = _why(it, kws, known or {})
                    target = dest
                    if not why:
                        if maybe_dir and _maybe(it):
                            target = Path(maybe_dir); target.mkdir(parents=True, exist_ok=True)
                            rep["maybe"] += 1
                        else:
                            rep["filtered"] += 1
                            continue
                    else:
                        rep["saved"] += 1
                        rep["why"][why] = rep["why"].get(why, 0) + 1
                    eml = self.call("GET", f"/users/{urllib.parse.quote(self.mailbox)}/messages/{m['id']}/$value", raw=True)
                    when = re.sub(r"[^0-9]", "", m.get("receivedDateTime", ""))[:12]
                    p = target / f"{when} {_safe(it.SenderName or it.SenderEmailAddress)} {_safe(it.Subject)} [{m['id'][-8:]}].eml"
                    p.write_bytes(eml)
                    seen.add(key)
                    if target == dest:
                        saved.append(p)
                except Exception as e:
                    rep["failed"].append(f"{m.get('subject')!r}: {e}")
            url = page.get("@odata.nextLink")
        return saved, seen, rep

    def reply_on_thread(self, subject, html, attachment, to=None, reply_all=True, folder="inbox"):
        """Create a draft reply on the newest mail of the thread (same subject), attach the file, return the draft's
        link (opens in Outlook). If no mail with that subject is found, a new draft with that subject is created.
        The draft is NEVER sent."""
        u = f"/users/{urllib.parse.quote(self.mailbox)}"
        key = re.sub(r"(?i)^((re|fw|fwd):\s*)+", "", subject or "").strip()
        esc = key.replace("'", "''")
        q = urllib.parse.urlencode({"$filter": f"contains(subject,'{esc}')", "$top": "25",
                                    "$select": "id,subject,receivedDateTime"})
        found = [m for m in self.call("GET", f"{u}/messages?{q}").get("value", [])
                 if re.sub(r"(?i)^((re|fw|fwd):\s*)+", "", m.get("subject") or "").strip().lower() == key.lower()]
        base = max(found, key=lambda m: m.get("receivedDateTime", "")) if found else None
        if base:
            d = self.call("POST", f"{u}/messages/{base['id']}/{'createReplyAll' if reply_all else 'createReply'}", {})
            cur = self.call("GET", f"{u}/messages/{d['id']}?$select=body")
            self.call("PATCH", f"{u}/messages/{d['id']}", {"body": {"contentType": "html",
                      "content": f"<p>{html}</p>" + (cur.get("body", {}).get("content") or "")}})
        else:
            d = self.call("POST", f"{u}/messages", {"subject": subject, "body": {"contentType": "html", "content": f"<p>{html}</p>"},
                          "toRecipients": [{"emailAddress": {"address": a.strip()}} for a in (to or "").split(";") if a.strip()]})
        att = Path(attachment)
        self.call("POST", f"{u}/messages/{d['id']}/attachments", {
            "@odata.type": "#microsoft.graph.fileAttachment", "name": att.name,
            "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "contentBytes": base64.b64encode(att.read_bytes()).decode()})
        link = self.call("GET", f"{u}/messages/{d['id']}?$select=webLink").get("webLink")
        return link, bool(base)


    # ---- a OneDrive / SharePoint folder, read by its sharing link -------------------------------------------------
    @staticmethod
    def share_id(link):
        """Graph's 'encoded sharing URL': u! + url-safe base64 of the link, without padding."""
        return "u!" + base64.urlsafe_b64encode(link.encode()).decode().rstrip("=")

    def sync_folder(self, link, dest, seen_state, exts=(".pdf", ".msg", ".eml", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".txt"),
                    max_files=5000):
        """Copy NEW or CHANGED files from the shared folder (any depth) into dest. Read only: nothing is changed in
        OneDrive. seen_state is {item_id: eTag}; returns (new_paths, seen_state, report)."""
        dest = Path(dest); dest.mkdir(parents=True, exist_ok=True)
        state = dict(seen_state)
        rep = dict(found=0, new=0, unchanged=0, skipped=0, failed=[])
        new = []
        root = self.call("GET", f"/shares/{self.share_id(link)}/driveItem?$select=id,name,parentReference,folder")
        drive = (root.get("parentReference") or {}).get("driveId")
        todo = [(root["id"], Path())]
        while todo and rep["found"] < max_files:
            item_id, rel = todo.pop()
            url = f"/drives/{drive}/items/{item_id}/children?$top=200&$select=id,name,eTag,size,folder,file"
            while url:
                page = self.call("GET", url)
                for it in page.get("value", []):
                    if "folder" in it:
                        todo.append((it["id"], rel / _safe(it["name"])))
                        continue
                    rep["found"] += 1
                    name = it["name"]
                    if not name.lower().endswith(exts):
                        rep["skipped"] += 1
                        continue
                    if state.get(it["id"]) == it.get("eTag"):
                        rep["unchanged"] += 1
                        continue
                    try:
                        data = self.call("GET", f"/drives/{drive}/items/{it['id']}/content", raw=True)
                        p = dest / rel / name
                        p.parent.mkdir(parents=True, exist_ok=True)
                        p.write_bytes(data)
                        state[it["id"]] = it.get("eTag")
                        new.append(p)
                        rep["new"] += 1
                    except Exception as e:
                        rep["failed"].append(f"{name}: {e}")
                url = page.get("@odata.nextLink")
        return new, state, rep
