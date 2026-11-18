"""Fake shared folder -> scan -> save through the web screens -> check the Excel he gets."""
import importlib
import os
import json
import sys

import openpyxl

sys.path.insert(0, str(__import__("pathlib").Path(__file__).resolve().parents[1] / "sample"))
import make_samples  # noqa: E402


def test_full_flow(tmp_path, monkeypatch):
    make_samples.main(tmp_path / "fake")
    cfg = json.load(open(__import__("pathlib").Path(__file__).resolve().parents[1] / "config.example.json"))
    cfg.update(shared_folder=str(tmp_path / "fake" / "shared_folder"), sources=[{"name": "Shared folder", "path": str(tmp_path / "fake" / "shared_folder")}],
               consolidated_list=str(tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"),
               output_folder=str(tmp_path / "out"), data_folder=str(tmp_path / "data"), write_mode="copy")
    (tmp_path / "config.json").write_text(json.dumps(cfg))
    monkeypatch.setenv("BC_CONFIG", str(tmp_path / "config.json"))
    web = importlib.reload(importlib.import_module("web"))
    c = web.app.test_client()

    assert c.post("/scan").status_code == 302
    items = [web.store.get(i) for i in range(1, 20) if web.store.get(i)]
    assert len(items) == 6, "7 files, one is a byte-identical copy and must be skipped"
    by_file = {i["path"].split("/")[-1]: i for i in items}
    custs = web.customers()
    first = next(v for v in items if "/North/" in v["path"] and v["path"].endswith(".pdf") and "scan" not in v["path"])
    assert first["suggested_status"] == "DISPUTE"
    assert first["suggested_theirs"] == round(custs[first["cust_no"]]["erp"] - 70000, 2)
    nochange = by_file["Re Balance confirmation.eml"]
    assert nochange["suggested_status"] == "CLEAR"
    scan = next(v for k, v in by_file.items() if k.endswith("scan.pdf"))
    assert scan["needs_typing"]
    caveat = by_file["balance conf.eml"]
    assert caveat["needs_typing"] and caveat["flags"]
    mismatch = next(v for k, v in by_file.items() if k.endswith("reply.pdf"))
    assert not mismatch["name_ok"]

    # Owner checks two: accepts the dispute, types the handwritten one.
    for it, theirs in [(first, first["suggested_theirs"]), (scan, custs[scan["cust_no"]]["erp"] - 1520.5)]:
        r = c.post(f"/save/{it['id']}", data=dict(cust_no=it["cust_no"], theirs=str(theirs), remarks="test"))
        assert r.status_code == 302
    for page in ["/", "/check", "/saved", "/mis"]:
        assert c.get(page).status_code == 200, page

    out = sorted((tmp_path / "out").glob("Balance_Confirmation_*.xlsx"))[-1]
    wb = openpyxl.load_workbook(out)
    ws = wb["Consolidated List "]
    row = custs[first["cust_no"]]["row"]
    assert ws.cell(row, 9).value == 70000.0           # DIFFERENCE
    assert ws.cell(row, 10).value == "Dispute"        # Clear / Dispute
    assert ws.cell(row, 19).value == "Received"       # Balance Confirmation Status
    row2 = custs[scan["cust_no"]]["row"]
    assert ws.cell(row2, 9).value == 1520.5
    mis = wb["MIS."]
    # Received is a live formula over his consolidated list; the tool does not type the number itself
    assert any(str(mis.cell(r, 7).value or "").startswith("=COUNTIFS(") for r in range(4, 40))
    assert mis.cell(4, 9).value.startswith("=")       # his formulas are kept
    page = c.get("/mis").data.decode()
    assert page.count("Email MIS to boss") == 1      # one e-mail button, not several
    # Outlook draft
    r = c.get("/mis-email")
    assert r.status_code == 200 and b"X-Unsent: 1" in r.data


def test_in_place_keeps_his_edits_and_waits_when_open(tmp_path, monkeypatch):
    make_samples.main(tmp_path / "fake")
    live = tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"
    cfg = json.load(open(__import__("pathlib").Path(__file__).resolve().parents[1] / "config.example.json"))
    cfg.update(shared_folder=str(tmp_path / "fake" / "shared_folder"), sources=[{"name": "Shared folder", "path": str(tmp_path / "fake" / "shared_folder")}], consolidated_list=str(live),
               output_folder=str(tmp_path / "out"), data_folder=str(tmp_path / "data"), write_mode="in_place")
    (tmp_path / "config.json").write_text(json.dumps(cfg))
    monkeypatch.setenv("BC_CONFIG", str(tmp_path / "config.json"))
    web = importlib.reload(importlib.import_module("web"))
    c = web.app.test_client()
    c.post("/scan")
    custs = web.customers()
    items = [web.store.get(i) for i in range(1, 20) if web.store.get(i)]
    a = next(i for i in items if i["suggested_status"] == "DISPUTE")
    b = next(i for i in items if i["suggested_status"] == "CLEAR")

    # He edits a column the tool does not own (Sales Person Name, col Q=17) in his own file.
    wb = openpyxl.load_workbook(live); ws = wb["Consolidated List "]
    ws.cell(custs[a["cust_no"]]["row"], 17, "Edited by owner"); wb.save(live)

    c.post(f"/save/{a['id']}", data=dict(cust_no=a["cust_no"], theirs=str(a["suggested_theirs"])))
    ws = openpyxl.load_workbook(live)["Consolidated List "]
    r = custs[a["cust_no"]]["row"]
    assert ws.cell(r, 10).value == "Dispute" and ws.cell(r, 17).value == "Edited by owner"
    assert list((tmp_path / "data" / "backups").glob("*.xlsx"))

    # Excel open -> nothing written, warning shown; closed -> written on next page.
    lock = live.parent / ("~$" + live.name); lock.write_text("x")
    c.post(f"/save/{b['id']}", data=dict(cust_no=b["cust_no"], theirs=str(b["suggested_theirs"])))
    rb = custs[b["cust_no"]]["row"]
    assert openpyxl.load_workbook(live)["Consolidated List "].cell(rb, 10).value is None
    assert b"Excel is open" in c.get("/").data
    lock.unlink()
    c.get("/saved")
    assert openpyxl.load_workbook(live)["Consolidated List "].cell(rb, 10).value == "Clear"
    assert b"Excel is open" not in c.get("/").data


def test_mail_and_folder_are_merged_without_duplicates(tmp_path, monkeypatch):
    import email
    from email import policy
    make_samples.main(tmp_path / "fake")
    shared = tmp_path / "fake" / "shared_folder"
    mail = tmp_path / "fake" / "his_mail"; mail.mkdir()

    # Reply A: arrives in his mail (.eml with a PDF) AND an RM saved just the PDF -> one item, seen in both.
    eml_path = next(shared.rglob("Re Balance confirmation.eml"))
    msg = email.message_from_bytes(eml_path.read_bytes(), policy=policy.default)
    pdf_bytes = next(msg.iter_attachments()).get_content()
    (shared / "North" / "RM saved this one.pdf").write_bytes(pdf_bytes)
    (mail / "Re Balance confirmation.eml").write_bytes(eml_path.read_bytes())
    eml_path.unlink()
    # Reply B: only in his mail (RM never saved it).
    only_mail = mail / "only mail.eml"
    only_mail.write_text("From: x@example.com\nSubject: Balance confirmation CN10005\n\nWe confirm, no changes.\n")

    cfg = json.load(open(__import__("pathlib").Path(__file__).resolve().parents[1] / "config.example.json"))
    cfg.update(sources=[{"name": "Shared folder", "path": str(shared)}, {"name": "His mail", "path": str(mail)}],
               consolidated_list=str(tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"),
               output_folder=str(tmp_path / "out"), data_folder=str(tmp_path / "data"), write_mode="copy")
    (tmp_path / "config.json").write_text(json.dumps(cfg))
    monkeypatch.setenv("BC_CONFIG", str(tmp_path / "config.json"))
    web = importlib.reload(importlib.import_module("web"))
    c = web.app.test_client()
    c.post("/scan"); c.post("/scan")  # twice: nothing may be added the second time
    items = web.store.all_items()
    both = [i for i in items if len(i["seen_in"]) == 2]
    assert len(both) == 1 and set(both[0]["seen_in"]) == {"Shared folder", "His mail"}
    mail_only = [i for i in items if i["seen_in"] == ["His mail"]]
    assert len(mail_only) == 1 and mail_only[0]["path"].endswith("only mail.eml")
    # This mail never states any balance: our balance must come from his consolidated list.
    m = mail_only[0]
    assert m["cust_no"] == "CN10005" and m["suggested_status"] == "CLEAR"
    assert m["suggested_theirs"] == web.customers()["CN10005"]["erp"] > 0
    page = c.get("/compare").data.decode()
    assert "in both" in page and "only in His mail" in page and "only mail.eml" in page


def test_boss_gets_mis_only(tmp_path, monkeypatch):
    import email
    from email import policy
    make_samples.main(tmp_path / "fake")
    cfg = json.load(open(__import__("pathlib").Path(__file__).resolve().parents[1] / "config.example.json"))
    cfg.update(sources=[{"name": "Shared folder", "path": str(tmp_path / "fake" / "shared_folder")}],
               consolidated_list=str(tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"),
               output_folder=str(tmp_path / "out"), data_folder=str(tmp_path / "data"), write_mode="copy",
               boss_email="boss@example.com")
    (tmp_path / "config.json").write_text(json.dumps(cfg))
    monkeypatch.setenv("BC_CONFIG", str(tmp_path / "config.json"))
    web = importlib.reload(importlib.import_module("web"))
    c = web.app.test_client()
    # Outlook draft
    msg = email.message_from_bytes(c.get("/mis-email").data, policy=policy.default)
    atts = list(msg.iter_attachments())
    assert len(atts) == 1 and atts[0].get_filename().startswith("MIS_")
    assert "consolidated" not in msg.get_body().get_content().lower()
    import io
    wb = openpyxl.load_workbook(io.BytesIO(atts[0].get_content()))
    assert [n.strip().lower() for n in wb.sheetnames] == ["mis."]
    # plain download + mailto point at the MIS file, not the full list
    wb2 = openpyxl.load_workbook(io.BytesIO(c.get("/download-mis").data))
    assert len(wb2.sheetnames) == 1
    page = c.get("/mis").data.decode()
    assert "Email MIS to boss" in page and "/open/mis" in page


def test_open_excel_button(tmp_path, monkeypatch):
    make_samples.main(tmp_path / "fake")
    cfg = json.load(open(__import__("pathlib").Path(__file__).resolve().parents[1] / "config.example.json"))
    cfg.update(sources=[{"name": "Shared folder", "path": str(tmp_path / "fake" / "shared_folder")}],
               consolidated_list=str(tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"),
               output_folder=str(tmp_path / "out"), data_folder=str(tmp_path / "data"), write_mode="in_place", host="127.0.0.1")
    (tmp_path / "config.json").write_text(json.dumps(cfg))
    monkeypatch.setenv("BC_CONFIG", str(tmp_path / "config.json"))
    web = importlib.reload(importlib.import_module("web"))
    opened = []
    monkeypatch.setattr(web, "open_with_default_app", lambda p: opened.append(str(p)))
    c = web.app.test_client()
    assert c.get("/open/list").status_code == 200
    assert opened[-1].endswith("Consolidated_List_FAKE.xlsx")        # his real file, in his own folder
    assert c.get("/open/mis").status_code == 200 and "MIS_Balance_Confirmation_" in opened[-1]
    assert c.get("/open/other").status_code == 404
    # tool on a server: nothing opens, he gets the location instead
    web.CFG["host"] = "0.0.0.0"
    n = len(opened)
    r = c.get("/open/list")
    assert len(opened) == n and b"not running on your own computer" in r.data


def test_login_required_when_password_set(tmp_path, monkeypatch):
    make_samples.main(tmp_path / "fake")
    cfg = json.load(open(__import__("pathlib").Path(__file__).resolve().parents[1] / "config.example.json"))
    cfg.update(sources=[{"name": "Shared folder", "path": str(tmp_path / "fake" / "shared_folder")}],
               consolidated_list=str(tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"),
               output_folder=str(tmp_path / "out"), data_folder=str(tmp_path / "data"), write_mode="in_place",
               host="127.0.0.1", access_password="s3cret")
    (tmp_path / "config.json").write_text(json.dumps(cfg))
    monkeypatch.setenv("BC_CONFIG", str(tmp_path / "config.json"))
    web = importlib.reload(importlib.import_module("web"))
    c = web.app.test_client()
    r = c.get("/mis")
    assert r.status_code == 302 and "/login" in r.headers["Location"]
    assert c.get("/health").status_code == 302 and c.get("/file/1").status_code == 302
    assert "Wrong password" in c.post("/login", data=dict(password="nope", who="x")).data.decode()
    assert c.post("/login", data=dict(password="s3cret", who="Owner")).status_code == 302
    assert c.get("/").status_code == 200


def test_visitor_through_ngrok_cannot_open_excel_on_host_pc(tmp_path, monkeypatch):
    make_samples.main(tmp_path / "fake")
    cfg = json.load(open(__import__("pathlib").Path(__file__).resolve().parents[1] / "config.example.json"))
    cfg.update(sources=[{"name": "Shared folder", "path": str(tmp_path / "fake" / "shared_folder")}],
               consolidated_list=str(tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"),
               output_folder=str(tmp_path / "out"), data_folder=str(tmp_path / "data"), write_mode="in_place", host="127.0.0.1")
    (tmp_path / "config.json").write_text(json.dumps(cfg))
    monkeypatch.setenv("BC_CONFIG", str(tmp_path / "config.json"))
    web = importlib.reload(importlib.import_module("web"))
    opened = []
    monkeypatch.setattr(web, "open_with_default_app", lambda p: opened.append(str(p)))
    c = web.app.test_client()
    r = c.get("/open/list", headers={"X-Forwarded-For": "1.2.3.4", "Host": "abc.ngrok-free.app"})
    assert r.status_code == 200 and opened == []                 # nothing opened on the host PC
    assert "Open this file in Excel" in r.data.decode()
    assert c.get("/open/list").status_code == 200 and len(opened) == 1   # local visit still opens it


def test_replies_for_customers_already_entered_by_hand_skip_the_queue(tmp_path, monkeypatch):
    make_samples.main(tmp_path / "fake")
    cfg = json.load(open(__import__("pathlib").Path(__file__).resolve().parents[1] / "config.example.json"))
    lst = tmp_path / "fake" / "Consolidated_List_FAKE.xlsx"
    cfg.update(sources=[{"name": "Shared folder", "path": str(tmp_path / "fake" / "shared_folder")}],
               consolidated_list=str(lst), output_folder=str(tmp_path / "out"), data_folder=str(tmp_path / "data"),
               write_mode="in_place", host="127.0.0.1")
    (tmp_path / "config.json").write_text(json.dumps(cfg))
    monkeypatch.setenv("BC_CONFIG", str(tmp_path / "config.json"))
    web = importlib.reload(importlib.import_module("web"))
    c = web.app.test_client()
    c.post("/scan")
    items = [i for i in web.store.new_items() if i["cust_no"]]
    assert items
    no = items[0]["cust_no"]
    total_before = len(web.new_items())
    wb = openpyxl.load_workbook(lst)
    from bc import workbook
    ws = workbook._sheet(wb, "consolidated"); h, cols = workbook._header(ws)
    row = web.customers()[no]["row"]
    import datetime as dt
    ws.cell(row, cols["recvd"], dt.date(2026, 11, 10)); ws.cell(row, cols["status"], "Clear")
    wb.save(lst)
    web.STATE["customers"] = {}
    assert len(web.new_items()) < total_before
    page = c.get("/already").data.decode()
    assert no in page and "Check anyway" in page
    assert c.get("/").status_code == 200
