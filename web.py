"""Balance confirmation review tool. Run:  python web.py   then open http://<server>:8080"""
import json
import os
from pathlib import Path

from flask import Flask, abort, redirect, render_template, request, send_file, url_for

from bc import checks, dump, ingest, outlook, rules, workbook
from bc.store import Store

HERE = Path(__file__).parent
CFG = json.loads((Path(os.environ.get("BC_CONFIG", HERE / "config.json"))).read_text())
DATA = Path(CFG["data_folder"]); DATA.mkdir(parents=True, exist_ok=True)
store = Store(DATA / "balance_confirmations.db")
app = Flask(__name__)
app.secret_key = (CFG.get("secret_key") or CFG.get("access_password") or "bc-local") + "|bc"

STATE = {"customers": {}, "unmapped": set()}
RENDER_LOCK = __import__("threading").Lock()   # PDFium is not safe to use from several threads at once


def sources():
    """Where replies come from. 'sources' in config, or the single 'shared_folder' (old config)."""
    base = CFG["sources"] if CFG.get("sources") else [{"name": "Shared folder", "path": CFG["shared_folder"]}]
    if onedrive_on():     # the OneDrive/SharePoint folder is copied here (read only) and read like any other folder
        base = list(base) + [{"name": "OneDrive folder", "path": str(DATA / "onedrive_mirror")}]
    return base


def scan_all():
    if onedrive_on() and not STATE.get("od_syncing"):
        STATE["od_syncing"] = True
        try:
            onedrive_sync_now()
        except Exception as e:
            app.logger.warning("OneDrive sync failed: %s", e)
        finally:
            STATE["od_syncing"] = False
    return ingest.scan(sources(), store, customers(), DATA, CFG.get("own_domains", ["gpglobal.com"]))


def customers():
    if not STATE["customers"]:
        STATE["customers"] = workbook.load_customers(CFG["consolidated_list"])
    return STATE["customers"]


def user():
    from flask import session
    return session.get("who") or request.headers.get("X-Remote-User") or CFG.get("user_name", "owner")


@app.before_request
def require_login():
    """If access_password is set in the config, nobody gets in without it (use this whenever the page can be
    reached from outside his own PC, e.g. through ngrok). Off when the password is empty."""
    from flask import session
    pw = CFG.get("access_password")
    if not pw or request.endpoint in ("login", "static"):
        return None
    if session.get("ok") is True:
        return None
    return redirect(url_for("login", next=request.path))


@app.route("/login", methods=["GET", "POST"])
def login():
    from flask import session
    import hmac, time
    err = ""
    if request.method == "POST":
        time.sleep(1)                                   # slows down guessing
        if hmac.compare_digest(request.form.get("password", ""), str(CFG.get("access_password", ""))):
            session["ok"] = True
            session["who"] = (request.form.get("who") or "owner").strip()[:40]
            nxt = request.args.get("next") or "/"
            return redirect(nxt if nxt.startswith("/") and not nxt.startswith("//") else "/")
        err = "Wrong password."
    return render_template("login.html", err=err)


def fmt(v):
    if v is None:
        return ""
    neg = v < 0
    v = abs(v)
    whole, frac = f"{v:.2f}".split(".")
    head, tail = whole[:-3], whole[-3:]
    groups = []
    while len(head) > 2:
        groups.insert(0, head[-2:]); head = head[:-2]
    if head:
        groups.insert(0, head)
    s = ",".join(groups + [tail]) if groups else tail
    return ("-" if neg else "") + "₹ " + s + "." + frac


app.jinja_env.filters["inr"] = fmt


def pre_done():
    """Customers whose Excel row already has a received date or a Clear/Dispute mark that the tool did not write:
    the owner entered those himself before. Their replies are not put in the 'to check' queue."""
    mine = {d["no"] for d in store.decisions()}
    return {no for no, c in customers().items() if (c.get("recvd") or c.get("status")) and no not in mine}


def new_items(skip_pre=True):
    items = store.new_items()
    if not skip_pre:
        return items
    pre = pre_done()
    return [it for it in items if it["cust_no"] not in pre]


@app.route("/already")
def already():
    pre, cs = pre_done(), customers()
    rows = [it for it in new_items(skip_pre=False) if it["cust_no"] in pre]
    return render_template("already.html", rows=rows, cs=cs)


@app.route("/")
def home():
    cs = customers()
    try:
        scan_all()  # pick up new files every time Home opens
    except OSError as e:
        app.logger.warning("shared folder not reachable: %s", e)
    pending = sum(1 for c in cs.values() if not c["recvd"]) - store.count("saved")
    return render_template("home.html", todo=len(new_items()), already=len(new_items(False)) - len(new_items()),
                           done=store.saved_today(),
                           pending=max(pending, 0), period=CFG.get("period", ""))


@app.post("/scan")
def scan():
    n = scan_all()
    return redirect(url_for("home", added=n))


@app.route("/check")
@app.route("/check/<int:item_id>")
def check(item_id=None):
    if item_id:
        it = store.get(item_id)
    else:
        after = int(request.args.get("after", 0))
        it = next((x for x in new_items() if x["id"] > after), None)
    if not it:
        return render_template("done.html")
    cs = customers()
    c = cs.get(it["cust_no"]) if it["cust_no"] else None
    earlier = store.saved_for_customer(it["cust_no"], it["id"]) if c else []
    total = len(new_items())
    vp = it["view_path"] or it["path"]
    pages = page_count(vp) if vp.lower().endswith(".pdf") else 0
    return render_template("check.html", it=it, c=c, earlier=earlier, total=total, pages=pages,
                           filename=Path(it["path"]).name,
                           customers=sorted(cs.values(), key=lambda x: x["name"]),
                           theirs=fmt(it["suggested_theirs"]).replace("₹ ", "") if it["suggested_theirs"] is not None else "")


@app.post("/save/<int:item_id>")
def save(item_id):
    it = store.get(item_id) or abort(404)
    no = request.form.get("cust_no", "").split(" · ")[0].strip()
    c = customers().get(no)
    if not c:
        return redirect(url_for("check", item_id=item_id, err="Pick a customer from the list."))
    theirs = rules.parse_amount(request.form.get("theirs", "").replace("₹", ""))
    if theirs is None:
        return redirect(url_for("check", item_id=item_id, err="Type the customer's balance."))
    status = rules.status_for(c["erp"], theirs)
    store.save(item_id, no, theirs, status, request.form.get("remarks", "").strip(), user())
    write()
    return redirect(url_for("check", after=0))


@app.post("/skip/<int:item_id>")
def skip(item_id):
    return redirect(url_for("check", after=item_id))


@app.post("/not-a-reply/<int:item_id>")
def not_a_reply(item_id):
    store.set_state(item_id, "not_a_reply", user())
    return redirect(url_for("check"))


@app.post("/reopen/<int:item_id>")
def reopen(item_id):
    store.reopen(item_id, user())
    return redirect(url_for("check", item_id=item_id))


@app.route("/file/<int:item_id>")
def file(item_id):
    it = store.get(item_id) or abort(404)
    p = it["view_path"] or it["path"]
    return send_file(p, mimetype=__import__("mimetypes").guess_type(p)[0] or "text/plain")


@app.route("/att/<int:item_id>/<int:n>")
def attachment(item_id, n):
    """Any attachment of a mail (the signed letter, a picture, a ledger) so he can look at it."""
    it = store.get(item_id) or abort(404)
    try:
        path = it["attachments"][n]["path"]
    except IndexError:
        abort(404)
    return send_file(path)


def page_count(path):
    try:
        import pypdfium2 as pdfium
        with RENDER_LOCK:
            return len(pdfium.PdfDocument(path))
    except Exception:
        return 0


def io_png(img):
    import io
    b = io.BytesIO()
    img.save(b, "PNG")
    b.seek(0)
    return b


@app.route("/diag/<int:item_id>")
def diag(item_id):
    """Plain-text report on why a file will not draw. He can paste it to me."""
    import importlib.metadata as md, traceback
    it = store.get(item_id) or abort(404)
    p = it["view_path"] or it["path"]
    lines = [f"file: {p}", f"exists: {os.path.exists(p)}", f"size: {os.path.getsize(p) if os.path.exists(p) else '-'}"]
    for pkg in ("pypdfium2", "Pillow", "Flask"):
        try:
            lines.append(f"{pkg}: {md.version(pkg)}")
        except Exception as e:
            lines.append(f"{pkg}: missing ({e})")
    try:
        import pypdfium2 as pdfium
        with RENDER_LOCK:
            pdf = pdfium.PdfDocument(p)
            lines.append(f"pages: {len(pdf)}")
            im = pdf[0].render(scale=1.6).to_pil().convert("L")
        lo, hi = im.getextrema()
        lines.append(f"render page 1: OK, size {im.size}, darkest pixel {lo} (255 = completely white = blank)")
        saved = DATA / "pages_v2" / f"{it['file_hash'][:16]}_0.png"
        lines.append(f"saved picture: {saved.name} {saved.stat().st_size if saved.exists() else 'none yet'} bytes")
    except Exception:
        lines.append(traceback.format_exc())
    return "\n".join(lines), 200, {"Content-Type": "text/plain; charset=utf-8"}


@app.route("/page/<int:item_id>/<int:n>.png")
def page(item_id, n):
    """One PDF page as a picture, so it shows the same in every browser."""
    it = store.get(item_id) or abort(404)
    out = DATA / "pages_v2" / f"{it['file_hash'][:16]}_{n}.png"
    if out.exists():   # an older version could leave a half-written picture behind: check it, redo if broken
        try:
            from PIL import Image
            with Image.open(out) as im:
                im.verify()
        except Exception:
            out.unlink(missing_ok=True)
    if not out.exists():
        try:
            import pypdfium2 as pdfium
            out.parent.mkdir(parents=True, exist_ok=True)
            with RENDER_LOCK:
                pdf = pdfium.PdfDocument(it["view_path"] or it["path"])
                img = pdf[n].render(scale=1.6).to_pil()
                if img.convert("L").getextrema()[0] >= 250:   # a completely white picture: never keep it
                    app.logger.warning("page %s of item %s drew blank", n, item_id)
                    return send_file(io_png(img), mimetype="image/png")
                tmp = out.with_suffix(".tmp.png")
                img.save(tmp)
                os.replace(tmp, out)
        except Exception as e:
            app.logger.exception("could not draw page %s of item %s", n, item_id)
            return (f"Could not draw this page: {type(e).__name__}: {e}", 500, {"Content-Type": "text/plain"})
    return send_file(out, mimetype="image/png")


@app.route("/compare")
def compare():
    """Which replies are in both his mail and the shared folder, and which are only in one."""
    names = [x["name"] for x in sources()]
    groups = {"both": [], "only": {n: [] for n in names}}
    for it in store.all_items():
        seen = [n for n in it["seen_in"] if n in names]
        if len(seen) > 1:
            groups["both"].append(it)
        elif seen:
            groups["only"][seen[0]].append(it)
    return render_template("compare.html", groups=groups, names=names, cs=customers())


@app.route("/saved")
def saved():
    return render_template("saved.html", rows=store.recent_saved(50), cs=customers())


WRITE_LOCK = __import__("threading").RLock()    # one writer at a time: two people saving together must not clash


def write():
    with WRITE_LOCK:
        return _write()


def _write():
    """Write his Excel. In 'in_place' mode his real file is updated (only the tool's columns);
    if he has it open, the write waits and is retried on the next page he opens."""
    if CFG.get("write_mode") == "in_place":
        try:
            path, missing = workbook.write_in_place(CFG["consolidated_list"], store.decisions(),
                                                    DATA / "backups", CFG["division_map"],
                                                    fill_counts=bool(CFG.get("fill_mis_counts")),
                                                    received_formula=CFG.get("mis_formulas", CFG.get("mis_received_formula", True)))
            STATE["pending_write"] = False
            STATE["customers"] = {}  # re-read: he may have edited other columns
            return path, missing
        except workbook.FileLocked:
            STATE["pending_write"] = True
            return Path(CFG["consolidated_list"]), []
    return workbook.write_outputs(CFG["consolidated_list"], store.decisions(), CFG["output_folder"],
                                  CFG["division_map"], fill_counts=bool(CFG.get("fill_mis_counts")),
                                  received_formula=CFG.get("mis_formulas", CFG.get("mis_received_formula", True)))


@app.before_request
def retry_pending_write():
    if STATE.get("pending_write") and request.endpoint not in ("page", "file", "static"):
        write()


@app.context_processor
def inject_pending():
    return {"pending_write": STATE.get("pending_write", False), "outlook_on": bool(CFG.get("outlook_folder")) and runs_on_his_computer(), "imap_on": imap_on() or graph_on(), "graph_on": graph_on(), "onedrive_on": onedrive_on(), "maybe_count": len(list(maybe_dir().glob("*.msg"))) if maybe_dir().exists() else 0, "pull_report": STATE.get("pull_report"),
            "outlook_mail_on": (bool(CFG.get("outlook_folder")) or __import__("sys").platform == "win32") and runs_on_his_computer()}


def run_checks():
    """Issues in his consolidated list right now (empty if the list can't be read)."""
    try:
        return checks.check_rows(checks.read_rows(CFG["consolidated_list"]), store.acked())
    except Exception as e:  # unreadable file must never stop the tool
        app.logger.warning("checks skipped: %s", e)
        return []


@app.context_processor
def inject_check_count():
    if request.endpoint in ("check", "page", "file", "static"):
        return {}
    return {"issue_count": len(run_checks())}


@app.route("/checks")
def checks_page():
    issues = run_checks()
    tie = []
    try:
        tie = checks.mis_tieout(CFG["consolidated_list"], CFG["division_map"])
    except Exception as e:
        app.logger.warning("MIS tie-out skipped: %s", e)
    return render_template("checks.html", issues=issues, tie=tie, done=request.args.get("done"))


@app.post("/checks/act")
def checks_act():
    no, code, action = (request.form.get(k, "") for k in ("no", "code", "action"))
    c = next((x for x in checks.read_rows(CFG["consolidated_list"]) if x["no"] == no), None)
    if not c:
        abort(404)
    if action == "accept":
        store.ack(no, code, user())
    elif action in ("recalc", "dispute", "clear") and c["theirs"] is not None and c["erp"] is not None:
        status = {"recalc": rules.status_for(c["erp"], c["theirs"]), "dispute": "DISPUTE", "clear": "CLEAR"}[action]
        store.manual_decision(no, c["theirs"], status, "Corrected on Checks page", user())
        write()
        STATE["customers"] = {}
    return redirect(url_for("checks_page", done=no))


@app.route("/dump", methods=["GET", "POST"])
def dump_page():
    """Start a new period: upload the ERP dump, get a new consolidated list in his template. His live file is never touched."""
    report = None
    if request.method == "POST":
        f = request.files.get("dump")
        if not f or not f.filename.lower().endswith(".xlsx"):
            return render_template("dump.html", report=None, err="Choose the ERP dump file (.xlsx).", teams=CFG.get("regional_teams", {}))
        up = DATA / "uploads"; up.mkdir(parents=True, exist_ok=True)
        src = up / f"dump_{__import__('datetime').datetime.now():%Y%m%d_%H%M%S}.xlsx"
        f.save(src)
        out = Path(CFG["output_folder"]) / f"Consolidated_List_NEW_{__import__('datetime').date.today():%Y-%m-%d}.xlsx"
        try:
            report = dump.build(src, CFG["consolidated_list"], out, regional_teams=CFG.get("regional_teams", {}))
        except (ValueError, StopIteration, KeyError) as e:
            return render_template("dump.html", report=None, err=f"I could not read this file: {e}", teams=CFG.get("regional_teams", {}))
    return render_template("dump.html", report=report, err=None, teams=CFG.get("regional_teams", {}))


@app.route("/dump/download")
def dump_download():
    files = sorted(Path(CFG["output_folder"]).glob("Consolidated_List_NEW_*.xlsx"))
    if not files:
        abort(404)
    return send_file(files[-1], as_attachment=True)


@app.route("/mis")
def mis():
    path, missing = write()
    import openpyxl
    rows = mis_rows(path)
    _, unmapped = workbook.mis_counts([dict(region=c["region"], group=c["group"], priority=c["priority"], received=False)
                                       for c in customers().values()], CFG["division_map"])
    workbook.write_mis_only(path, CFG["output_folder"], CFG["division_map"])
    return render_template("mis.html", rows=rows, unmapped=sorted(unmapped), missing=missing,
                           boss=CFG.get("boss_email", ""), period=CFG.get("period", ""),
                           mailto=mailto_link(rows, path))


def mis_rows(path):
    import openpyxl
    rows = []
    ws = workbook._sheet(openpyxl.load_workbook(path), "mis")
    # The numbers shown here are counted from the consolidated list (the same counts his MIS formulas make), so the
    # page and the mail summary never depend on Excel having calculated the formulas yet.
    cust = workbook.load_customers(path).values()
    counts, _ = workbook.mis_counts(
        [dict(region=c["region"], group=c["group"], priority=c["priority"],
              received=bool(c.get("recvd")) or str(c.get("bc_status") or "").lower() == "received") for c in cust],
        CFG["division_map"])
    region = None
    for r in range(4, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a and "total" in str(a).lower():
            region = None
            continue
        if a:
            region = str(a)
        if region and b:
            t_ = counts.get((region.strip().upper(), workbook.ALIASES.get(str(b).strip().upper(), str(b).strip().upper())))
            if not t_ or not t_["total"]:
                continue
            t, p1, p3, rec = t_["total"], t_["p1"], t_["p3"], t_["received"]
            p2 = t - p1 - p3
            rows.append(dict(region=region, div=b, total=t, p1=p1, p2=p2, p3=p3, rec=rec, pen=t - p3 - rec,
                             pct=f"{(rec / (p1 + p2) * 100):.0f}%" if (p1 + p2) else "-"))
    return rows


def summary_lines(rows):
    totals = {}
    for r in rows:
        t = totals.setdefault(r["region"], [0, 0, 0])
        t[0] += r["total"] or 0; t[1] += r["rec"] or 0; t[2] += r["pen"] or 0
    return [f"{reg}: {rec} received, {pen} pending (of {tot})" for reg, (tot, rec, pen) in totals.items()]


def mailto_link(rows, path):
    path = Path(CFG["output_folder"]) / f"MIS_Balance_Confirmation_{__import__('datetime').date.today():%Y-%m-%d}.xlsx"
    """A mailto: link. Clicking it opens a new Outlook message with To, Subject and the summary filled in.
    mailto cannot attach files, so the body gives the report's location on the shared drive instead."""
    import datetime as dt
    from urllib.parse import quote
    totals = {}
    for r in rows:
        t = totals.setdefault(r["region"], [0, 0, 0])
        t[0] += r["total"] or 0; t[1] += r["rec"] or 0; t[2] += r["pen"] or 0
    lines = [f"{reg}: {rec} received, {pen} pending (of {tot})" for reg, (tot, rec, pen) in totals.items()]
    shown = CFG.get("report_link_prefix", "") + Path(path).name if CFG.get("report_link_prefix") else str(path)
    body = ("Dear Sir,\n\nBalance confirmation status as on " + f"{dt.date.today():%d %b %Y}:\n\n" + "\n".join(lines) +
            "\n\nMIS report:\n" + shown + "\n\nRegards")
    subject = f"Balance confirmation MIS - {dt.date.today():%d %b %Y}"
    return f"mailto:{quote(CFG.get('boss_email', ''))}?subject={quote(subject)}&body={quote(body)}"


@app.route("/download")
def download():
    path, _ = write()
    return send_file(path, as_attachment=True)


def mis_file():
    """Write his Excel, then make the MIS-only file for the boss. Returns its path."""
    path, _ = write()
    return workbook.write_mis_only(path, CFG["output_folder"], CFG["division_map"])


def runs_on_his_computer():
    """True when the tool runs on the same computer as his browser (then it can open Excel for him)."""
    if not (CFG.get("host", "0.0.0.0") in ("127.0.0.1", "localhost") or CFG.get("open_on_server")):
        return False
    try:    # a visitor coming through ngrok / a proxy is NOT sitting at this computer
        if request.headers.get("X-Forwarded-For") or request.headers.get("X-Forwarded-Host") or \
                request.host.split(":")[0] not in ("127.0.0.1", "localhost", "[::1]"):
            return False
    except RuntimeError:
        pass
    return True


def open_with_default_app(path):
    import subprocess
    import sys
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    elif sys.platform.startswith("win"):
        os.startfile(str(path))  # noqa: S606 (Windows only)
    else:
        subprocess.Popen(["xdg-open", str(path)])


@app.route("/open/<what>")
def open_excel(what):
    """Open his Excel (the consolidated list, or the MIS file for the boss) in Excel on his computer,
    in the same folder as the real file, so he can edit it there. Only works when the tool runs on his own computer."""
    if what == "list":
        path, _ = write()
        path = Path(path)
    elif what == "mis":
        path = Path(mis_file())
    else:
        abort(404)
    if not runs_on_his_computer():
        return render_template("opened.html", path=str(path), opened=False, what=what, ok=False)
    try:
        open_with_default_app(path)
        return render_template("opened.html", path=str(path), opened=True, what=what, ok=True)
    except Exception as e:  # no default app, permissions, ...
        app.logger.warning("could not open %s: %s", path, e)
        return render_template("opened.html", path=str(path), opened=False, what=what, ok=False)


@app.route("/download-mis")
def download_mis():
    return send_file(mis_file(), as_attachment=True)


def maybe_dir():
    return Path(CFG.get("outlook_maybe_folder") or DATA / "maybe")


def maybe_items():
    d = maybe_dir()
    if not d.exists():
        return []
    out = []
    for p in sorted(d.glob("*.msg"), reverse=True):
        try:
            info = ingest.read_msg(p, DATA / "maybe_tmp")
        except Exception:
            info = dict(mail_from="", mail_subject=p.name, text="", attachments=[])
        out.append(dict(name=p.name, sender=info.get("mail_from", ""), subject=info.get("mail_subject", "") or p.name,
                        preview=(info.get("text", "") or "")[:300],
                        atts=[a["name"] for a in info.get("attachments", [])]))
    return out


@app.route("/maybe")
def maybe():
    return render_template("maybe.html", items=maybe_items())


@app.post("/maybe/<action>")
def maybe_act(action):
    import shutil
    d = maybe_dir()
    name = Path(request.form.get("name", "")).name          # a bare file name only
    src = d / name
    if name and src.exists() and action in ("accept", "dismiss"):
        if action == "accept":
            target = Path(sources()[0]["path"])
            shutil.move(str(src), str(target / name))
            scan_all()
        else:
            (d / "_dismissed").mkdir(exist_ok=True)
            shutil.move(str(src), str(d / "_dismissed" / name))
    return redirect(url_for("maybe"))


def onedrive_on():
    return bool(CFG.get("graph_tenant") and CFG.get("graph_client_id") and CFG.get("onedrive_link"))


def onedrive_sync_now():
    import json as _j
    state = DATA / "onedrive_state.json"
    st = _j.loads(state.read_text()) if state.exists() else {}
    new, st, rep = graph_client().sync_folder(CFG["onedrive_link"], DATA / "onedrive_mirror", st)
    state.write_text(_j.dumps(st))
    STATE["onedrive_report"] = rep
    return rep


@app.post("/onedrive/sync")
def onedrive_sync_route():
    try:
        onedrive_sync_now()
    except Exception as e:
        return (f"<p style='font-family:sans-serif'>Could not read the OneDrive folder: {e}</p><p><a href='/'>Back</a></p>", 200)
    n = scan_all()
    return redirect(url_for("home", added=n))


def graph_on():
    return bool(CFG.get("graph_tenant") and CFG.get("graph_client_id") and CFG.get("graph_mailbox"))


def graph_client():
    from bc import graph
    secret = os.environ.get("BC_GRAPH_SECRET") or CFG.get("graph_secret", "")
    return graph.Graph(CFG["graph_tenant"], CFG["graph_client_id"], secret, CFG.get("graph_mailbox", ""))


def graph_pull_now():
    import json as _j
    state = DATA / "graph_seen.json"
    seen = set(_j.loads(state.read_text())) if state.exists() else set()
    dest = Path(CFG.get("graph_save_to") or sources()[0]["path"])
    saved, seen, rep = graph_client().pull(dest, seen, keywords=CFG.get("outlook_keywords", []), known=known_senders(),
                                           maybe_dir=maybe_dir(), folder=CFG.get("graph_folder", "inbox"),
                                           since_days=int(CFG.get("graph_since_days", 30)))
    state.write_text(_j.dumps(sorted(seen)))
    STATE["pull_report"] = rep
    return rep


def imap_on():
    return bool(CFG.get("imap_host") and CFG.get("imap_user"))


def imap_pull_now():
    """Pull new relevant mail from the IMAP mailbox into the watched folder. The password comes from the
    environment (BC_IMAP_PASSWORD), never from the config file."""
    import json as _j
    from bc import imap
    state = DATA / "imap_seen.json"
    seen = set(_j.loads(state.read_text())) if state.exists() else set()
    dest = Path(CFG.get("imap_save_to") or sources()[0]["path"])
    pw = os.environ.get("BC_IMAP_PASSWORD") or CFG.get("imap_password", "")
    saved, seen, rep = imap.pull(CFG["imap_host"], CFG["imap_user"], pw, CFG.get("imap_folder", "INBOX"), dest, seen,
                                 keywords=CFG.get("outlook_keywords", []), known=known_senders(), maybe_dir=maybe_dir(),
                                 port=int(CFG.get("imap_port", 993)), ssl=CFG.get("imap_ssl", True),
                                 since_days=int(CFG.get("imap_since_days", 30)))
    state.write_text(_j.dumps(sorted(seen)))
    STATE["pull_report"] = rep
    return rep


@app.post("/imap/pull")
def imap_pull_route():
    try:
        if graph_on():
            graph_pull_now()
        else:
            imap_pull_now()
    except Exception as e:
        return (f"<p style='font-family:sans-serif'>Could not read the mailbox: {e}</p><p><a href='/'>Back</a></p>", 200)
    n = scan_all()
    return redirect(url_for("home", added=n))


def known_senders():
    """Addresses, company domains and numbers of the customers in his list, to recognise relevant mail."""
    from bc.rules import EMAIL_RE
    emails, domains, numbers = set(), set(), set()
    for no, c in customers().items():
        numbers.add(str(no).upper())
        for a in EMAIL_RE.findall(str(c.get("email") or "").lower()):
            emails.add(a)
            domains.add(a.split("@")[-1])
    return dict(emails=emails, domains=domains, numbers=numbers)


@app.route("/outlook/pull", methods=["POST"])
def outlook_pull():
    """Classic Outlook on his PC: copy new mails from one Outlook folder into the watched folder, then scan."""
    import json as _j
    state = DATA / "outlook_seen.json"
    seen = set(_j.loads(state.read_text())) if state.exists() else set()
    dest = Path(CFG.get("outlook_save_to") or sources()[0]["path"])
    try:
        saved, seen, rep = outlook.pull(CFG.get("outlook_folders") or CFG.get("outlook_folder", "Inbox"), dest, seen, keywords=CFG.get("outlook_keywords", []), known=known_senders(), maybe_dir=maybe_dir())
    except Exception as e:
        return (f"<p style='font-family:sans-serif'>Could not read Outlook: {e}</p><p><a href='/'>Back</a></p>", 200)
    state.write_text(_j.dumps(sorted(seen)))
    n = scan_all()
    STATE["pull_report"] = rep
    return redirect(url_for("home", added=n))


@app.route("/mis-email")
def mis_email():
    if request.args.get("graph") == "1" and graph_on():
        import datetime as dt
        path = mis_file()
        live, _ = write()
        try:
            link, found = graph_client().reply_on_thread(
                CFG.get("thread_subject") or f"Balance confirmation MIS - {dt.date.today():%d %b %Y}",
                "Dear Sir,<br><br>Balance confirmation status as on " + f"{dt.date.today():%d %b %Y}:<br>"
                + "<br>".join(summary_lines(mis_rows(live))) + "<br><br>Please find the MIS attached.<br><br>Regards",
                path, to=CFG.get("thread_to") or CFG.get("boss_email", ""))
        except Exception as e:
            return (f"<p style='font-family:sans-serif'>Could not create the draft: {e}</p><p><a href='/mis'>Back</a></p>", 200)
        return redirect(link) if link else redirect(url_for("mis"))
    if request.args.get("outlook") == "1":
        path = mis_file()
        live, _ = write()
        import datetime as dt
        try:
          outlook.reply_draft(CFG.get("thread_subject") or f"Balance confirmation MIS - {dt.date.today():%d %b %Y}",
                                 CFG.get("boss_email", ""), path,
                                 "Dear Sir,<br><br>Balance confirmation status as on " + f"{dt.date.today():%d %b %Y}:<br>"
                                 + "<br>".join(summary_lines(mis_rows(live))) +
                                 "<br><br>Please find the MIS attached.<br><br>Regards",
                                 CFG.get("outlook_thread_folder"), reply_all=CFG.get("outlook_reply_all", True),
                                 search=bool(CFG.get("thread_subject")))
        except Exception as e:
            return (f"<p style='font-family:sans-serif'>Could not open Outlook: {e}</p><p><a href='/mis'>Back</a></p>", 200)
        return redirect(url_for("mis"))
    """An Outlook draft (.eml with X-Unsent) with the MIS report (only) attached. He reviews and clicks Send."""
    import datetime as dt
    from email.message import EmailMessage
    path = mis_file()
    m = EmailMessage()
    m["To"] = CFG.get("thread_to") or CFG.get("boss_email", "")
    if CFG.get("thread_cc"):
        m["Cc"] = CFG["thread_cc"]
    ts = (CFG.get("thread_subject") or "").strip()
    # Same subject as the existing daily thread, so Outlook files it in that conversation.
    m["Subject"] = (ts if ts.lower().startswith("re:") else f"RE: {ts}") if ts \
        else f"Balance confirmation MIS - {dt.date.today():%d %b %Y}"
    mid = (CFG.get("thread_message_id") or "").strip()
    if mid:                                   # optional: Message-ID of the last mail in the thread
        m["In-Reply-To"] = mid
        m["References"] = mid
    m["X-Unsent"] = "1"
    m.set_content("Dear Sir,\n\nPlease find attached today's balance confirmation MIS.\n\nRegards")
    m.add_attachment(Path(path).read_bytes(), maintype="application",
                     subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=Path(path).name)
    out = DATA / "MIS_email.eml"
    out.write_bytes(bytes(m))
    return send_file(out, as_attachment=True, download_name="MIS_email_to_boss.eml")


@app.route("/health")
def health():
    """Go-live check: can the tool see everything it needs? Open /health after installing on the server."""
    import openpyxl
    res = {}
    def ok(name, fn):
        try:
            res[name] = dict(ok=True, detail=fn())
        except Exception as e:
            res[name] = dict(ok=False, detail=f"{type(e).__name__}: {e}")
    lst = Path(CFG["consolidated_list"])
    ok("consolidated list exists", lambda: str(lst) if lst.exists() else (_ for _ in ()).throw(FileNotFoundError(str(lst))))
    def read_list():
        c = workbook.load_customers(lst)
        return f"{len(c)} customers read"
    ok("consolidated list readable (sheet and headers)", read_list)
    def mis_sheet():
        wb = openpyxl.load_workbook(lst)
        return f"sheets: {wb.sheetnames}; MIS sheet: {workbook._sheet(wb, 'mis').title!r}"
    ok("MIS sheet found", mis_sheet)
    for s in sources():
        ok(f"reply folder: {s['name']}", lambda s=s: f"{len(list(Path(s['path']).rglob('*')))} files" if Path(s["path"]).exists() else (_ for _ in ()).throw(FileNotFoundError(s["path"])))
    def writable(d):
        d = Path(d); d.mkdir(parents=True, exist_ok=True)
        t = d / ".bc_write_test"; t.write_text("x"); t.unlink()
        return f"{d} is writable"
    ok("reports folder writable", lambda: writable(CFG["output_folder"]))
    ok("data folder writable", lambda: writable(CFG["data_folder"]))
    ok("list not locked by Excel", lambda: "free" if not workbook.is_open_in_excel(lst) else (_ for _ in ()).throw(RuntimeError("Excel has it open")))
    ok("mode", lambda: f"write_mode={CFG.get('write_mode')}, host={CFG.get('host')}, runs on his computer={runs_on_his_computer()}")
    res["all_ok"] = all(v["ok"] for k, v in res.items())
    return json.dumps(res, indent=1), 200, {"Content-Type": "application/json"}


def _auto_scan():
    import threading, time
    mins = float(CFG.get("auto_scan_minutes") or 0)
    if mins <= 0:
        return
    def loop():
        while True:
            time.sleep(mins * 60)
            try:
                if onedrive_on():
                    onedrive_sync_now()
                if graph_on():
                    graph_pull_now()
                elif imap_on():
                    imap_pull_now()
                scan_all()
            except Exception as e:      # never let a bad file stop the timer
                app.logger.warning("auto scan failed: %s", e)
    threading.Thread(target=loop, daemon=True).start()


if __name__ == "__main__":
    _auto_scan()
    host, port = CFG.get("host", "0.0.0.0"), int(CFG.get("port", 8080))
    try:                                   # a real server (works on Windows too); falls back to Flask's own
        from waitress import serve
        print(f"Serving on http://{host}:{port}  (waitress)")
        serve(app, host=host, port=port, threads=8)
    except ImportError:
        app.run(host=host, port=port)
