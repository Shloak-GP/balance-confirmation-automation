"""Read replies straight from a mailbox over IMAP (no Outlook, no M365 needed): works on a server.
Needs IT to allow IMAP for one mailbox (or a mailbox that receives the replies). Mail is only READ:
the folder is opened read-only and messages are fetched with BODY.PEEK, so nothing is marked read, moved or deleted.
Uses the same relevance test as the Outlook pull (customer address/domain/number, attachment name, our letter, keyword)."""
import email
import email.policy
import imaplib
import re
from email.utils import parseaddr
from pathlib import Path

from .outlook import _maybe, _safe, _why


class _Mail:
    """Makes an e-mail look like an Outlook item, so the same rules apply."""

    def __init__(self, msg):
        self.Subject = str(msg.get("subject", "") or "")
        self.SenderName, self.SenderEmailAddress = parseaddr(str(msg.get("from", "") or ""))
        body = msg.get_body(preferencelist=("plain", "html"))
        try:
            self.Body = body.get_content() if body else ""
        except Exception:
            self.Body = ""
        self.Attachments = [type("A", (), {"FileName": a.get_filename() or ""})() for a in msg.iter_attachments()]
        self.Class = 43


def pull(host, user, password, folder, dest, seen_ids, keywords=(), known=None, maybe_dir=None,
         port=993, ssl=True, since_days=30, limit=2000, client=None):
    """Returns (saved_paths, seen_ids, report). seen_ids holds 'folder|uidvalidity|uid' so a mail is fetched once."""
    dest = Path(dest)
    dest.mkdir(parents=True, exist_ok=True)
    seen = set(seen_ids)
    kws = [k.strip().lower() for k in (keywords or []) if k and k.strip()]
    rep = dict(in_folder=0, saved=0, already=0, not_mail=0, filtered=0, maybe=0, why={}, failed=[])
    saved = []
    own = client is None
    M = client or (imaplib.IMAP4_SSL(host, port) if ssl else imaplib.IMAP4(host, port))
    try:
        if own:
            M.login(user, password)
        typ, _ = M.select(f'"{folder}"', readonly=True)
        if typ != "OK":
            raise RuntimeError(f"Could not open the mailbox folder {folder!r}")
        import datetime as dt
        since = (dt.date.today() - dt.timedelta(days=since_days)).strftime("%d-%b-%Y")
        typ, data = M.uid("search", None, "SINCE", since)
        uids = (data[0] or b"").split()[-limit:]
        for uid in uids:
            rep["in_folder"] += 1
            key = f"{folder}|{uid.decode()}"
            if key in seen:
                rep["already"] += 1
                continue
            try:
                typ, parts = M.uid("fetch", uid, "(BODY.PEEK[])")
                raw = next(p[1] for p in parts if isinstance(p, tuple))
                msg = email.message_from_bytes(raw, policy=email.policy.default)
                it = _Mail(msg)
                why = "all"
                if kws or known:
                    why = _why(it, kws, known or {})
                target = dest
                if not why:
                    if maybe_dir and _maybe(it):
                        target = Path(maybe_dir)
                        target.mkdir(parents=True, exist_ok=True)
                        rep["maybe"] += 1
                    else:
                        rep["filtered"] += 1
                        continue
                else:
                    rep["saved"] += 1
                    rep["why"][why] = rep["why"].get(why, 0) + 1
                name = f"{(msg.get('date') and re.sub(r'[^0-9]', '', str(msg.get('date')))[:12]) or 'mail'} " \
                       f"{_safe(it.SenderName or it.SenderEmailAddress)} {_safe(it.Subject)} [{uid.decode()}].eml"
                (target / name[:150]).write_bytes(raw)
                seen.add(key)
                if target == dest:
                    saved.append(target / name[:150])
            except Exception as e:
                rep["failed"].append(f"uid {uid.decode()}: {e}")
    finally:
        if own:
            try:
                M.logout()
            except Exception:
                pass
    return saved, seen, rep
