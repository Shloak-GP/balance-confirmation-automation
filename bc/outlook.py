"""Classic Outlook on Windows, through Outlook's own local automation (pywin32). No M365, no server, no IT.

Two jobs:
  pull(folder_path, dest, seen_ids, app=None)  save new mails of one Outlook folder as .msg into the watched folder
  reply_draft(subject, to, attachment, body, app=None)  open a reply on the existing thread with the MIS attached

`app` is injectable so the logic is tested without Outlook. It only runs on a Windows PC with CLASSIC Outlook
(the 'new Outlook' has no such interface).
"""
import re
from pathlib import Path

OL_MSG = 3  # olMSG: save as .msg


def _app():
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError("Needs Windows with classic Outlook and 'pip install pywin32'.") from e
    pythoncom.CoInitialize()      # every web-server thread must do this once before using COM
    return win32com.client.Dispatch("Outlook.Application")


def _release():
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:
        pass


def find_folder(app, path):
    """'Inbox/Balance Confirmation' (or 'Mailbox - Name/Inbox/...') -> the Outlook folder."""
    ns = app.GetNamespace("MAPI")
    parts = [p for p in re.split(r"[\\/]+", path) if p]
    if not parts:
        raise ValueError("outlook_folder is empty")
    if parts[0].lower() == "inbox":
        cur = ns.GetDefaultFolder(6)           # olFolderInbox
        parts = parts[1:]
    else:
        cur = ns.Folders[parts[0]]
        parts = parts[1:]
    for p in parts:
        cur = cur.Folders[p]
    return cur


def _safe(s):
    return re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", s or "").strip()[:60] or "mail"


def pull(folder_path, dest, seen_ids, app=None, limit=None, keywords=(), known=None, maybe_dir=None):
    """folder_path may be one Outlook folder path or a list of them (e.g. his Inbox and a 'Balance Confirmation' folder)."""
    if isinstance(folder_path, (list, tuple)):
        own = app is None
        app = app or _app()
        try:
            tot = dict(in_folder=0, saved=0, already=0, not_mail=0, filtered=0, maybe=0, why={}, failed=[])
            allsaved, seen = [], set(seen_ids)
            for fp in folder_path:
                sv, seen, r = _pull(app, fp, dest, seen, limit, keywords, known, maybe_dir)
                allsaved += sv
                for k in ("in_folder", "saved", "already", "not_mail", "filtered", "maybe"):
                    tot[k] += r[k]
                for k, v in r["why"].items():
                    tot["why"][k] = tot["why"].get(k, 0) + v
                tot["failed"] += r["failed"]
            return allsaved, seen, tot
        finally:
            if own:
                _release()
    return _pull_one(folder_path, dest, seen_ids, app, limit, keywords, known, maybe_dir)


def _pull_one(folder_path, dest, seen_ids, app=None, limit=None, keywords=(), known=None, maybe_dir=None):
    """Save mails not seen before as .msg. Returns (saved_paths, seen_ids, report). Mail in Outlook is only read."""
    own = app is None
    app = app or _app()
    try:
        return _pull(app, folder_path, dest, seen_ids, limit, keywords, known, maybe_dir)
    finally:
        if own:
            _release()


FREE_MAIL = {"gmail.com", "yahoo.com", "yahoo.co.in", "yahoo.in", "rediffmail.com", "hotmail.com", "outlook.com",
             "live.com", "icloud.com", "aol.com"}
ATT_HINT = ("confirm", "balance", "ledger", "statement", "outstanding", "reconcil", "gp petro")


LETTER_PHRASES = ("confirmation of balance", "in connection with the audit of our accounts",
                  "balance confirmation", "confirm the balance", "balance as per our books",
                  "gp petroleums", "erp.noreply")


def _why(it, kws, known):
    """Why this mail looks relevant, as a short label ('' = no reason, leave it in Outlook). Any ONE reason is enough,
    so a reply that never says 'balance' is still caught:
      sender      the sender's address is a customer's address in the consolidated list
      domain      the sender's company domain (not Gmail/Yahoo) is a customer's
      number      a customer number (C#####) from the list is in the subject, body start or attachment names
      attachment  an attachment is named like a confirmation, ledger or statement
      our letter  our own letter's wording is in the mail (customers reply on it or quote it)
      keyword     a word from outlook_keywords is in the subject or first lines"""
    subj = str(getattr(it, "Subject", "") or "")
    body = str(getattr(it, "Body", "") or "")[:6000]
    names = []
    try:
        names = [str(a.FileName) for a in it.Attachments]
    except Exception:
        pass
    hay = f"{subj} {body} {' '.join(names)}"
    low = hay.lower()
    sender = str(getattr(it, "SenderEmailAddress", "") or "").lower()
    if sender and sender in known.get("emails", ()):
        return "sender"
    dom = sender.split("@")[-1] if "@" in sender else ""
    if dom and dom not in FREE_MAIL and dom in known.get("domains", ()):
        return "domain"
    nums = known.get("numbers", ())
    if nums and any(n in nums for n in re.findall(r"\bC\d{4,6}\b", hay.upper())):
        return "number"
    if any(h in n.lower() for n in names for h in ATT_HINT):
        return "attachment"
    if any(p in low for p in LETTER_PHRASES):
        return "our letter"
    if any(k in low for k in kws):
        return "keyword"
    return ""


ATT_EXT = (".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".xls", ".xlsx", ".doc", ".docx", ".heic")


def _maybe(it):
    """No clear reason to keep it, but it could be a reply: it has a document/picture attached, or is a RE:/FW:.
    These go to a separate 'maybe' folder for the owner to accept or dismiss. Obvious noise is not offered."""
    subj = str(getattr(it, "Subject", "") or "").strip().lower()
    if any(w in subj for w in ("unsubscribe", "newsletter", "webinar", "out of office", "automatic reply", "undeliverable")):
        return False
    try:
        if any(str(a.FileName).lower().endswith(ATT_EXT) for a in it.Attachments):
            return True
    except Exception:
        pass
    return subj.startswith(("re:", "fw:", "fwd:"))


def _pull(app, folder_path, dest, seen_ids, limit, keywords=(), known=None, maybe_dir=None):
    """Walks the folder AND its sub-folders. Nothing is skipped silently: the report counts what was saved,
    what was not an ordinary mail, and what failed (with the reason)."""
    root = find_folder(app, folder_path)
    dest = Path(dest)
    dest.mkdir(parents=True, exist_ok=True)
    seen = set(seen_ids)
    saved = []
    rep = dict(in_folder=0, saved=0, already=0, not_mail=0, filtered=0, maybe=0, why={}, failed=[])
    kws = [k.strip().lower() for k in (keywords or []) if k and k.strip()]

    def walk(folder):
        yield folder
        try:
            subs = list(folder.Folders)
        except Exception:
            subs = []
        for f in subs:
            yield from walk(f)

    for folder in walk(root):
        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass
        for it in items:
            rep["in_folder"] += 1
            if limit and rep["in_folder"] > limit:
                rep["failed"].append(f"stopped at the limit of {limit} items")
                return saved, seen, rep
            try:
                if getattr(it, "Class", 43) != 43:                 # 43 = an ordinary mail item
                    rep["not_mail"] += 1
                    continue
                eid = it.EntryID
                if eid in seen:
                    rep["already"] += 1
                    continue
                why = "all"
                if kws or known:
                    why = _why(it, kws, known or {})
                    if not why:
                        if maybe_dir and _maybe(it):
                            eid = it.EntryID
                            if eid in seen:
                                rep["already"] += 1
                                continue
                            md = Path(maybe_dir); md.mkdir(parents=True, exist_ok=True)
                            when = it.ReceivedTime.strftime("%Y%m%d_%H%M")
                            p = md / f"{when} {_safe(it.SenderName)} {_safe(it.Subject)} [{eid[-8:]}].msg"
                            if len(str(p)) > 240:
                                p = md / f"{when} [{eid[-12:]}].msg"
                            it.SaveAs(str(p), OL_MSG)
                            seen.add(eid)
                            rep["maybe"] += 1
                            continue
                        rep["filtered"] += 1      # left in Outlook, untouched, and looked at again on the next pull
                        continue
                when = it.ReceivedTime.strftime("%Y%m%d_%H%M")
                name = f"{when} {_safe(it.SenderName)} {_safe(it.Subject)} [{eid[-8:]}].msg"
                p = dest / name
                if len(str(p)) > 240:                              # Windows refuses paths over ~260 characters
                    p = dest / f"{when} [{eid[-12:]}].msg"
                it.SaveAs(str(p), OL_MSG)
                seen.add(eid)
                saved.append(p)
                rep["saved"] += 1
                rep["why"][why] = rep["why"].get(why, 0) + 1
            except Exception as e:
                rep["failed"].append(f"{getattr(it, 'Subject', '?')!r}: {e}")   # stays unseen, so it is retried
    return saved, seen, rep


def reply_draft(subject, to, attachment, body, folder_path=None, app=None, reply_all=True, search=True):
    """Find the latest mail of the thread (same subject, ignoring RE:/FW:) and open a reply with the file attached.
    If no mail is found, open a new mail with that subject. The mail is only DISPLAYED, never sent."""
    own = app is None
    app = app or _app()
    try:
        return _reply_draft(app, subject, to, attachment, body, folder_path, reply_all, search)
    finally:
        if own:
            _release()


def _reply_draft(app, subject, to, attachment, body, folder_path, reply_all, search):
    key = re.sub(r"(?i)^((re|fw|fwd):\s*)+", "", subject or "").strip().lower()
    base = None
    for fp in ((folder_path, "Inbox") if search else ()):
        if not fp:
            continue
        try:
            for it in find_folder(app, fp).Items:
                if re.sub(r"(?i)^((re|fw|fwd):\s*)+", "", getattr(it, "Subject", "") or "").strip().lower() == key:
                    if base is None or it.ReceivedTime > base.ReceivedTime:
                        base = it
        except Exception:
            pass
        if base is not None:
            break
    if base is not None:
        m = base.ReplyAll() if reply_all else base.Reply()   # reply-all keeps the thread's To and CC list
        m.HTMLBody = f"<p>{body}</p>" + (m.HTMLBody or "")
    else:
        m = app.CreateItem(0)
        m.Subject = subject
        m.HTMLBody = f"<p>{body}</p>"
    if to and not (getattr(m, "To", "") or "").strip():
        m.To = to
    m.Attachments.Add(str(Path(attachment).resolve()))
    m.Display()
    return base is not None
