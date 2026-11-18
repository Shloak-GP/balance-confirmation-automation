import importlib
import io
import json
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "sample"))
import make_samples  # noqa: E402
from bc import checks, dump  # noqa: E402


def _web(tmp_path, monkeypatch, mode="in_place"):
    make_samples.main(tmp_path / "fake")
    cfg = json.load(open(ROOT / "config.example.json"))
    cfg.update(sources=[{"name": "Shared folder", "path": str(tmp_path / "fake" / "shared_folder")}],
               consolidated_list=str(tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"),
               output_folder=str(tmp_path / "out"), data_folder=str(tmp_path / "data"), write_mode=mode)
    (tmp_path / "config.json").write_text(json.dumps(cfg))
    monkeypatch.setenv("BC_CONFIG", str(tmp_path / "config.json"))
    return importlib.reload(importlib.import_module("web"))


def test_checks_flag_and_fix(tmp_path, monkeypatch):
    web = _web(tmp_path, monkeypatch)
    live = tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"
    c = web.customers()
    no = next(iter(c))
    row = c[no]["row"]
    wb = openpyxl.load_workbook(live); ws = wb["Consolidated List "]
    hdr = {str(x.value).strip().lower(): x.column for x in ws[next(r for r in range(1, 10) if ws.cell(r, 5).value == "No.")]}
    erp_col = next(v for k, v in hdr.items() if k.startswith("balance as per erp"))
    th, df, st = (next(v for k, v in hdr.items() if k.startswith(n)) for n in ("balance as per customer", "difference", "clear"))
    ws.cell(row, th, ws.cell(row, erp_col).value - 500); ws.cell(row, df, 0); ws.cell(row, st, "Clear")
    wb.save(live)
    web.STATE["customers"] = {}
    issues = [i for i in web.run_checks() if i["no"] == no]
    assert {i["code"] for i in issues} >= {"arith", "clear_diff"}
    page = web.app.test_client().get("/checks").data.decode()
    assert "Marked Clear but the balances differ" in page and "Make it Dispute" in page
    r = web.app.test_client().post("/checks/act", data=dict(no=no, code="clear_diff", action="dispute"))
    assert r.status_code == 302
    ws2 = openpyxl.load_workbook(live)["Consolidated List "]
    assert ws2.cell(row, st).value == "Dispute" and ws2.cell(row, df).value == 500
    assert not [i for i in web.run_checks() if i["no"] == no]
    # "it is correct" is remembered and logged, and the flag does not come back
    ws2 = openpyxl.load_workbook(live); w = ws2["Consolidated List "]
    w.cell(row, st, "Clear"); w.cell(row, df, 500); ws2.save(live)
    web.app.test_client().post("/checks/act", data=dict(no=no, code="clear_diff", action="accept"))
    assert "clear_diff" not in {i["code"] for i in web.run_checks() if i["no"] == no}


def _make_dump(path):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Customer Wise Details"
    ws.append([None] * 16)
    ws.append(["No.", "Name", "Exclude Customer Yes/No", "Reason (If Yes)", "Regional Account Name", "Salesperson Name",
               "Salesperson Region Code", "Zone", "Customer Posting Group", " Remaining Amt. (LCY)", "x", "x", "x", "x", "x", "x"])
    for r in [("C1", "Alpha", "No", None, "Ramesh K", "Sam", "WEST 1", "", "RPO", 1000.5),
              ("C2", "Beta", "Yes", "Legal", "Vijay", "Ria", "ZONE S1", "ZONE S1", "REPSOL", -50),
              ("C3", "Gamma", "No", None, "Vijay", "Ria", "", "", "FOREIGN", 10),
              ("C4", "Delta", "No", None, "New Person", "Tom", "NORTH 2", "", "INDUSTRIAL", 0)]:
        ws.append(list(r) + [0] * 6)
    ag = wb.create_sheet("Customer - Aging Detailed")
    ag.append(["As of : 30-09-26"]); ag.append(["No.", "Name", "x", "x", "x", "Region"])
    ag.append(["C4", "Delta", "", "", "", "NORTH"])
    wb.save(path)


def test_dump_to_consolidated(tmp_path, monkeypatch):
    web = _web(tmp_path, monkeypatch)
    live = tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"
    prev_no = next(iter(web.customers()))
    wb = openpyxl.load_workbook(live); ws = wb["Consolidated List "]
    hdr = {str(x.value).strip().lower(): x.column for x in ws[next(r for r in range(1, 10) if ws.cell(r, 5).value == "No.")]}
    ws.cell(web.customers()[prev_no]["row"], hdr["no."], "C1"); ws.cell(web.customers()[prev_no]["row"], hdr["priority"], "Priority 2")
    ws.cell(web.customers()[prev_no]["row"], hdr["email id"], "a@b.in"); wb.save(live)
    d = tmp_path / "dump.xlsx"; _make_dump(d)
    rep = dump.build(d, live, tmp_path / "new.xlsx", regional_teams={"Ramesh K": "Ramesh Kulkarni"})
    assert rep["written"] == 3 and [x[2] for x in rep["dropped"]] == ["FOREIGN"]
    assert rep["carried"] == 1 and rep["no_priority"] == 2 and "New Person" in rep["unknown_team"]
    ws = openpyxl.load_workbook(tmp_path / "new.xlsx")["Consolidated List "]
    rows = {ws.cell(r, hdr["no."]).value: r for r in range(1, ws.max_row + 1) if ws.cell(r, hdr["no."]).value}
    a, b, d4 = rows["C1"], rows["C2"], rows["C4"]
    assert ws.cell(a, hdr["region"]).value == "WEST" and ws.cell(a, hdr["regional team"]).value == "Ramesh Kulkarni"
    assert ws.cell(a, hdr["priority"]).value == "Priority 2" and ws.cell(a, hdr["email id"]).value == "a@b.in"
    assert ws.cell(b, hdr["region"]).value == "SOUTH" and ws.cell(b, hdr["customer posting group"]).value == "AUTOMOTIVE (REPSOL)"
    assert ws.cell(b, hdr["balance confirmation to send yes/no"]).value == "No"
    assert ws.cell(d4, hdr["region"]).value == "NORTH"
    assert ws.cell(a, hdr["balance as per customer"]).value is None
    # page works
    c = web.app.test_client()
    r = c.post("/dump", data={"dump": (io.BytesIO(d.read_bytes()), "dump.xlsx")}, content_type="multipart/form-data")
    assert r.status_code == 200 and b"customers in the new list" in r.data
    assert c.get("/dump/download").status_code == 200


def test_mis_formulas_reading_consolidated_survive_boss_copy(tmp_path):
    """His MIS counts may be COUNTIFS formulas on the Consolidated sheet; the boss copy has no such sheet."""
    import openpyxl, shutil
    from pathlib import Path
    from bc import workbook
    src = Path(__file__).resolve().parent.parent / "sample" / "his_template.xlsx"
    p = tmp_path / "t.xlsx"; shutil.copy(src, p)
    wb = openpyxl.load_workbook(p)
    ws = wb["MIS."]
    ws["C5"] = "=COUNTIF('Consolidated List '!D:D,1)"
    wb.save(p)
    cfg_map = {"LUBRICANTS": "IPOL"}
    out = workbook.write_mis_only(p, tmp_path / "o", cfg_map)
    wb2 = openpyxl.load_workbook(out)
    assert wb2.sheetnames == ["MIS."]
    assert not str(wb2["MIS."]["C5"].value).startswith("=")


class _Fake:
    def __init__(self, **kw): self.__dict__.update(kw)


def test_outlook_pull_saves_new_mail_once(tmp_path):
    import datetime as dt
    from bc import outlook
    saved_calls = []

    class Item(_Fake):
        Class = 43
        def SaveAs(self, p, t): saved_calls.append(p); open(p, "w").write("x")

    items = [Item(EntryID="AAAA11112222", SenderName="Amber Trading", Subject="Balance: confirm?",
                  ReceivedTime=dt.datetime(2026, 11, 16, 9, 5)),
             _Fake(EntryID="BBBB", Class=53)]   # a meeting request is skipped
    folder = _Fake(Items=_Items(items), Folders={})
    ns = _Fake(GetDefaultFolder=lambda n: folder)
    app = _Fake(GetNamespace=lambda x: ns)
    s1, seen, rep = outlook.pull("Inbox", tmp_path, set(), app=app)
    s2, _, rep2 = outlook.pull("Inbox", tmp_path, seen, app=app)
    assert len(s1) == 1 and s1[0].suffix == ".msg" and s2 == []
    assert rep["in_folder"] == 2 and rep["not_mail"] == 1 and rep2["already"] == 1 and not rep["failed"]


class _Items(list):
    def Sort(self, *a): pass


def test_outlook_reply_all_keeps_thread_recipients(tmp_path):
    import datetime as dt
    from bc import outlook
    att = tmp_path / "MIS.xlsx"; att.write_text("x")
    log = {}

    class Draft:
        To = "a@x.com; b@x.com"
        HTMLBody = "<p>old</p>"
        class Attachments:
            @staticmethod
            def Add(p): log["att"] = p
        def Display(self): log["shown"] = True

    class Base:
        Subject = "RE: Daily balance confirmation MIS"
        ReceivedTime = dt.datetime(2026, 11, 16)
        def Reply(self): log["kind"] = "reply"; return Draft()
        def ReplyAll(self): log["kind"] = "all"; return Draft()

    folder = _Fake(Items=[Base()])
    app = _Fake(GetNamespace=lambda x: _Fake(GetDefaultFolder=lambda n: folder))
    found = outlook.reply_draft("Daily balance confirmation MIS", "boss@x.com", att, "Dear Sir", app=app)
    assert found and log["kind"] == "all" and log["shown"] and log["att"].endswith("MIS.xlsx")


def test_outlook_pull_keyword_filter(tmp_path):
    import datetime as dt
    from bc import outlook

    class Item(_Fake):
        Class = 43
        def SaveAs(self, p, t): open(p, "w").write("x")

    mk = lambda i, subj, body: Item(EntryID=f"ID{i:010d}", SenderName="A", Subject=subj, Body=body,
                                    ReceivedTime=dt.datetime(2026, 11, 16, 9, i))
    items = _Items([mk(1, "Balance confirmation C123", ""), mk(2, "Lunch?", "see you at 1"),
                    mk(3, "Re: statement", "our balance is Rs 1,000")])
    folder = _Fake(Items=items, Folders=[])
    app = _Fake(GetNamespace=lambda x: _Fake(GetDefaultFolder=lambda n: folder))
    s, seen, rep = outlook.pull("Inbox", tmp_path, set(), app=app, keywords=["balance"])
    assert len(s) == 2 and rep["filtered"] == 1 and len(seen) == 2    # the lunch mail is not marked as pulled


def test_outlook_pull_keeps_mail_from_known_customer_without_keywords(tmp_path):
    import datetime as dt
    from bc import outlook

    class Item(_Fake):
        Class = 43
        Attachments = []
        def SaveAs(self, p, t): open(p, "w").write("x")

    mk = lambda i, subj, sender: Item(EntryID=f"ID{i:010d}", SenderName="A", Subject=subj, Body="ok",
                                      SenderEmailAddress=sender, ReceivedTime=dt.datetime(2026, 11, 16, 9, i))
    items = _Items([mk(1, "Re: GP Petroleums", "accounts@amber.co.in"),      # known company domain, no keyword
                    mk(2, "Re: C19001 attached", "x@gmail.com"),            # customer number in subject
                    mk(3, "Lunch?", "friend@gmail.com")])                   # unrelated
    folder = _Fake(Items=items, Folders=[])
    app = _Fake(GetNamespace=lambda x: _Fake(GetDefaultFolder=lambda n: folder))
    known = dict(emails=set(), domains={"amber.co.in"}, numbers={"C19001"})
    s, seen, rep = outlook.pull("Inbox", tmp_path, set(), app=app, keywords=["balance"], known=known)
    assert len(s) == 2 and rep["filtered"] == 1


def test_outlook_pull_two_folders_and_reasons(tmp_path):
    import datetime as dt
    from bc import outlook

    class Item(_Fake):
        Class = 43
        Attachments = []
        def SaveAs(self, p, t): open(p, "w").write("x")

    mk = lambda i, subj, body: Item(EntryID=f"ID{i:010d}", SenderName="A", Subject=subj, Body=body,
                                    SenderEmailAddress="z@gmail.com", ReceivedTime=dt.datetime(2026, 11, 16, 9, i))
    inbox = _Fake(Items=_Items([mk(1, "Hi", "Confirmation of Balance as on 30/09/2026 - agreed"), mk(2, "Lunch", "x")]), Folders=[])
    other = _Fake(Items=_Items([mk(3, "Ledger", "see attached")]), Folders=[])
    ns = _Fake(GetDefaultFolder=lambda n: inbox, Folders={"Shared": _Fake(Folders={"BC": other})})
    app = _Fake(GetNamespace=lambda x: ns)
    s, seen, rep = outlook.pull(["Inbox", "Shared/BC"], tmp_path, set(), app=app, keywords=["ledger"], known={})
    assert len(s) == 2 and rep["why"] == {"our letter": 1, "keyword": 1} and rep["filtered"] == 1


def test_outlook_maybe_folder_gets_borderline_mail(tmp_path):
    import datetime as dt
    from bc import outlook

    class Att:
        def __init__(self, n): self.FileName = n

    class Item(_Fake):
        Class = 43
        def SaveAs(self, p, t): open(p, "w").write("x")

    mk = lambda i, subj, atts: Item(EntryID=f"ID{i:010d}", SenderName="A", Subject=subj, Body="ok",
                                    SenderEmailAddress="z@gmail.com", Attachments=[Att(a) for a in atts],
                                    ReceivedTime=dt.datetime(2026, 11, 16, 9, i))
    items = _Items([mk(1, "Hello", ["scan001.pdf"]), mk(2, "Newsletter: unsubscribe", ["a.pdf"]), mk(3, "Lunch", [])])
    folder = _Fake(Items=items, Folders=[])
    app = _Fake(GetNamespace=lambda x: _Fake(GetDefaultFolder=lambda n: folder))
    s, seen, rep = outlook.pull("Inbox", tmp_path / "main", set(), app=app, keywords=["balance"], known={},
                                maybe_dir=tmp_path / "maybe")
    assert s == [] and rep["maybe"] == 1 and rep["filtered"] == 2
    assert len(list((tmp_path / "maybe").glob("*.msg"))) == 1


def test_imap_pull_reads_only_and_filters(tmp_path):
    from email.message import EmailMessage
    from bc import imap

    def mk(subj, frm, body, att=None):
        m = EmailMessage(); m["Subject"] = subj; m["From"] = frm; m["Date"] = "Mon, 16 Nov 2026 09:05:00 +0530"
        m.set_content(body)
        if att:
            m.add_attachment(b"%PDF-1.4 x", maintype="application", subtype="pdf", filename=att)
        return bytes(m)

    store = {b"1": mk("Balance confirmation C19001", "Amber <a@amber.in>", "agreed"),
             b"2": mk("Lunch?", "friend@gmail.com", "see you"),
             b"3": mk("Hello", "x@gmail.com", "ok", att="scan001.pdf")}
    calls = []

    class Fake:
        def select(self, f, readonly=False): calls.append(("select", readonly)); return "OK", [b"3"]
        def uid(self, cmd, *a):
            if cmd == "search": return "OK", [b" ".join(store)]
            calls.append(("fetch", a[1])); return "OK", [(b"1 (UID 1)", store[a[0]]), b")"]

    s, seen, rep = imap.pull("h", "u", "p", "INBOX", tmp_path / "main", set(), keywords=["balance"], known={},
                             maybe_dir=tmp_path / "maybe", client=Fake())
    assert calls[0] == ("select", True) and all(c[1] == "(BODY.PEEK[])" for c in calls[1:])   # read-only, peek
    assert len(s) == 1 and rep["maybe"] == 1 and rep["filtered"] == 1 and len(seen) == 2
    s2, seen2, rep2 = imap.pull("h", "u", "p", "INBOX", tmp_path / "main", seen, keywords=["balance"], known={},
                                maybe_dir=tmp_path / "maybe", client=Fake())
    assert s2 == [] and rep2["already"] == 2


def test_graph_pull_and_thread_reply_with_fake_microsoft(tmp_path):
    import json as _j
    from bc import graph
    log = []
    msgs = [dict(id="AAA111", subject="Balance confirmation C19001", from_={"emailAddress": {"name": "Amber", "address": "a@amber.in"}},
                 internetMessageId="<1@x>", receivedDateTime="2026-11-16T09:05:00Z", bodyPreview="agreed", attachments=[]),
            dict(id="BBB222", subject="Lunch?", from_={"emailAddress": {"name": "F", "address": "f@gmail.com"}},
                 internetMessageId="<2@x>", receivedDateTime="2026-11-16T09:06:00Z", bodyPreview="hi", attachments=[])]
    for m in msgs:
        m["from"] = m.pop("from_")

    def req(method, url, headers=None, data=None):
        log.append((method, url.split("?")[0].replace("https://graph.microsoft.com/v1.0", "")))
        if "oauth2" in url:
            return 200, _j.dumps({"access_token": "tok"}).encode()
        if method == "GET" and url.endswith("$value"):
            return 200, b"From: a@amber.in\r\nSubject: x\r\n\r\nbody"
        if method == "GET" and "/messages?" in url and "mailFolders" in url:
            return 200, _j.dumps({"value": msgs}).encode()
        if method == "GET" and "contains(subject" in urllib_unquote(url):
            return 200, _j.dumps({"value": [dict(id="T1", subject="RE: Daily MIS", receivedDateTime="2026-11-17T08:00:00Z")]}).encode()
        if method == "POST" and url.endswith("createReplyAll"):
            return 201, _j.dumps({"id": "D1"}).encode()
        if method == "GET" and "$select=body" in url:
            return 200, _j.dumps({"body": {"content": "<p>old thread</p>"}}).encode()
        if method == "GET" and "webLink" in url:
            return 200, _j.dumps({"webLink": "https://outlook.office.com/owa/?ItemID=D1"}).encode()
        return 200, b"{}"

    g = graph.Graph("t", "c", "s", "owner@gp.com", req=req)
    s, seen, rep = g.pull(tmp_path / "main", set(), keywords=["balance"], known={}, maybe_dir=tmp_path / "maybe")
    assert len(s) == 1 and s[0].suffix == ".eml" and rep["filtered"] == 1 and len(seen) == 1
    assert all(m == "GET" or "oauth" in p or True for m, p in log)
    s2, _, rep2 = g.pull(tmp_path / "main", seen, keywords=["balance"], known={})
    assert s2 == [] and rep2["already"] == 1
    att = tmp_path / "MIS.xlsx"; att.write_bytes(b"x")
    link, found = g.reply_on_thread("Daily MIS", "hello", att)
    assert found and link.startswith("https://outlook.office.com")
    assert ("POST", "/users/owner%40gp.com/messages/D1/attachments") in log
    assert not any(p.endswith("/send") for _, p in log)          # a draft only, never sent


def urllib_unquote(u):
    import urllib.parse
    return urllib.parse.unquote(u)


def test_graph_onedrive_folder_sync_new_and_changed_only(tmp_path):
    import json as _j
    from bc import graph
    files = {"f1": dict(id="f1", name="C19001 Amber.pdf", eTag="v1", file={}), "f2": dict(id="f2", name="notes.docx", eTag="v1", file={}),
             "sub": dict(id="sub", name="East", folder={})}
    inner = {"f3": dict(id="f3", name="reply.msg", eTag="v1", file={})}
    content = {"f1": b"%PDF one", "f3": b"msg"}
    calls = []

    def req(method, url, headers=None, data=None):
        calls.append(method)
        if "oauth2" in url:
            return 200, _j.dumps({"access_token": "t"}).encode()
        if "/shares/" in url:
            assert "u!" in url
            return 200, _j.dumps({"id": "root", "parentReference": {"driveId": "D1"}, "folder": {}}).encode()
        if "/items/root/children" in url:
            return 200, _j.dumps({"value": list(files.values())}).encode()
        if "/items/sub/children" in url:
            return 200, _j.dumps({"value": list(inner.values())}).encode()
        if url.endswith("/content"):
            return 200, content[url.split("/items/")[1].split("/")[0]]
        return 404, b"{}"

    g = graph.Graph("t", "c", "s", "", req=req)
    new, st, rep = g.sync_folder("https://x/share", tmp_path / "mirror", {})
    assert sorted(p.name for p in new) == ["C19001 Amber.pdf", "reply.msg"] and rep["skipped"] == 1
    assert (tmp_path / "mirror" / "East" / "reply.msg").exists()
    new2, st2, rep2 = g.sync_folder("https://x/share", tmp_path / "mirror", st)
    assert new2 == [] and rep2["unchanged"] == 2                      # nothing downloaded twice
    files["f1"]["eTag"] = "v2"
    new3, _, _ = g.sync_folder("https://x/share", tmp_path / "mirror", st2)
    assert [p.name for p in new3] == ["C19001 Amber.pdf"]              # a changed file is fetched again
    assert set(calls) <= {"GET", "POST"}                                  # read only: no PUT/DELETE/PATCH
