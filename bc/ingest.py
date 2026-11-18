"""Pick up new reply files from the shared folder. Read-only on the folder: nothing is moved or edited."""
import email
import hashlib
from email import policy
from pathlib import Path

from pypdf import PdfReader

from . import rules

TYPES = {".pdf", ".eml", ".msg", ".txt"}


def file_hash(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def norm_key(*parts):
    """Fingerprint of the *content* of a reply, so the same reply is recognised however it arrived
    (RM saved the PDF, or it sits in his mail as .eml/.msg with that PDF attached)."""
    t = " ".join(" ".join(p.split()).lower() for p in parts if p)
    return hashlib.sha256(t.encode("utf-8", "ignore")).hexdigest()


def keep_attachment(name, size):
    """Skip inline signature pictures (image001.png, a few KB); keep everything a person might need to look at."""
    import re
    generic = re.fullmatch(r"(?i)image\d{0,3}\.(png|jpe?g|gif)", name or "")
    return not (generic and size < 30000) and size > 0


def save_attachment(attach_dir, name, data, used):
    attach_dir.mkdir(parents=True, exist_ok=True)
    safe = "".join(ch for ch in Path(name).name if ch not in '\\/:*?"<>|') or "attachment"
    out = attach_dir / safe
    n = 1
    while str(out) in used:
        n += 1
        out = attach_dir / f"{Path(safe).stem}_{n}{Path(safe).suffix}"
    out.write_bytes(data)
    used.add(str(out))
    return str(out)


def pick_view(atts):
    """The file to show next to the form: first PDF, else first picture."""
    for kind in ((".pdf",), (".png", ".jpg", ".jpeg", ".gif")):
        for a in atts:
            if a["name"].lower().endswith(kind):
                return a["path"]
    return None


def pdf_text(path_or_bytes):
    """Text layer only. Scanned pages give '' and go to the owner to type."""
    import io
    src = io.BytesIO(path_or_bytes) if isinstance(path_or_bytes, bytes) else str(path_or_bytes)
    try:
        r = PdfReader(src)
        return "\n".join((pg.extract_text() or "") for pg in r.pages)
    except Exception:
        return ""


def read_eml(path, attach_dir):
    msg = email.message_from_bytes(Path(path).read_bytes(), policy=policy.default)
    body = msg.get_body(preferencelist=("plain", "html"))
    text = body.get_content() if body else ""
    atts, used, key = [], set(), None
    for part in msg.iter_attachments():
        name = part.get_filename() or "attachment"
        data = part.get_content()
        if not isinstance(data, bytes) or not keep_attachment(name, len(data)):
            continue
        atts.append(dict(name=name, path=save_attachment(attach_dir, name, data, used)))
        if name.lower().endswith(".pdf"):
            key = key or hashlib.sha256(data).hexdigest()
            text += rules.ATT_MARK + pdf_text(data)
    key = key or norm_key(str(msg["subject"] or ""), body.get_content() if body else "")
    return dict(text=text, mail_from=str(msg["from"] or ""), mail_subject=str(msg["subject"] or ""),
                mail_date=str(msg["date"] or ""), view_path=pick_view(atts), attachments=atts, content_key=key)


def read_msg(path, attach_dir):
    """Outlook .msg, read with olefile (pure Python, no build step): body, subject, sender, every attachment."""
    try:
        import olefile
    except ImportError:
        return dict(text="", mail_from="", mail_subject=Path(path).name, mail_date="", view_path=None,
                    attachments=[], content_key=file_hash(path))
    o = olefile.OleFileIO(str(path))

    def prop(prefix, tag):
        for suffix, enc in (("001F", "utf-16-le"), ("001E", "latin-1")):
            n = f"{prefix}__substg1.0_{tag}{suffix}"
            if o.exists(n):
                return o.openstream(n).read().decode(enc, errors="ignore").rstrip("\x00")
        return ""

    body, subject, sender = prop("", "1000"), prop("", "0037"), prop("", "0C1A")
    text, atts, used, key = body, [], set(), None
    for d in o.listdir():
        if d[0].startswith("__attach_version1.0_") and len(d) == 2 and d[1].startswith("__substg1.0_3701"):
            pre = d[0] + "/"
            name = prop(pre, "3707") or prop(pre, "3704") or "attachment"
            data = o.openstream("/".join(d)).read()
            if not keep_attachment(name, len(data)):
                continue
            atts.append(dict(name=name, path=save_attachment(attach_dir, name, data, used)))
            if name.lower().endswith(".pdf"):
                key = key or hashlib.sha256(data).hexdigest()
                text += rules.ATT_MARK + pdf_text(data)
    o.close()
    key = key or norm_key(subject, body)
    return dict(text=text, mail_from=sender, mail_subject=subject, mail_date="", view_path=pick_view(atts),
                attachments=atts, content_key=key)


def read_file(p, attach_dir):
    ext = p.suffix.lower()
    if ext == ".pdf":
        return dict(text=pdf_text(p), mail_from="", mail_subject="", mail_date="", view_path=str(p), attachments=[],
                    content_key=file_hash(p))
    if ext == ".eml":
        return read_eml(p, attach_dir)
    if ext == ".msg":
        return read_msg(p, attach_dir)
    t = p.read_text(errors="ignore")
    return dict(text=t, mail_from="", mail_subject="", mail_date="", view_path=None, attachments=[], content_key=norm_key(t))


def scan(sources, store, customers, data_dir, own_domains=("gpglobal.com",)):
    """Read new files from every source. sources: [{"name": "Shared folder", "path": "..."}, ...].

    Read-only on every source. The same reply arriving twice (two RMs, or once in a source and once
    in his mail) becomes ONE item that remembers every source it was seen in.
    Returns number of new items.
    """
    if isinstance(sources, (str, Path)):
        sources = [{"name": "Shared folder", "path": str(sources)}]
    added = 0
    for src in sources:
        root = Path(src["path"])
        if not root.exists():
            continue
        for p in sorted(root.rglob("*")):
            if not p.is_file() or p.suffix.lower() not in TYPES or p.name.startswith("~$"):
                continue
            h = file_hash(p)
            info = read_file(p, Path(data_dir) / "attachments" / h[:12])
            key = info["content_key"]
            if store.has_key(key):
                store.add_seen(key, src["name"])
                continue
            full = " ".join([info["mail_subject"], info["text"]])
            no, how, name_ok, cands = rules.find_customer(full, p.name, customers, info["mail_from"], own_domains)
            if no is None and cands:
                hit = rules.pick_by_amount(info["text"], cands, customers)
                if hit:
                    no, how, name_ok = hit, "matched by name and the balance they state", True
            erp = customers[no]["erp"] if no else None
            s = rules.suggest(info["text"], erp, customers[no]["name"]) if no else dict(
                theirs=None, status=None, needs_typing=True, flags=[],
                reason="Customer not recognised. Please pick the customer.")
            store.add(path=str(p), view_path=info["view_path"], file_hash=h, content_key=key, seen_in=[src["name"]],
                      text=info["text"][:20000], mail_from=info["mail_from"], mail_subject=info["mail_subject"],
                      mail_date=info["mail_date"], attachments=info.get("attachments", []), cust_no=no, cust_how=how, name_ok=int(bool(name_ok)),
                      candidates=cands, suggested_theirs=s["theirs"], suggested_status=s["status"],
                      needs_typing=int(s["needs_typing"]), reason=s["reason"], flags=list(s["flags"]))
            added += 1
    return added
