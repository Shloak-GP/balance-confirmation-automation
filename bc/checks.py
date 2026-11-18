"""Self-checks on the consolidated list. They only FLAG; nothing is changed without his click.

Each issue: dict(code, no, name, row, level('fix'|'look'), message, options=[(action, label)]).
Actions are applied through the tool's own store (so they are logged) - see web.py /checks/act.
"""
import datetime as dt

import openpyxl

from . import rules, workbook

TOL = 0.005
BIG_SHARE = 0.5  # |difference| above half of the balance: usually a period total or a typo


def _d(v):
    if isinstance(v, dt.datetime):
        return v.date()
    return v if isinstance(v, dt.date) else None


def read_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = workbook._sheet(wb, "consolidated")
    hrow, cols = workbook._header(ws)
    rows = []
    for r in range(hrow + 1, ws.max_row + 1):
        no = ws.cell(r, cols["no"]).value
        if no in (None, ""):
            continue
        g = lambda k: ws.cell(r, cols[k]).value if k in cols else None
        theirs_raw = g("theirs")
        rows.append(dict(row=r, no=str(no).strip(), name=str(g("name") or "").strip(), erp=workbook._num(g("erp")),
                         theirs=workbook._num(theirs_raw), theirs_raw=theirs_raw, diff=workbook._num(g("diff")),
                         status=str(g("status") or "").strip().lower(), bc_status=str(g("bc_status") or "").strip().lower(),
                         sent=_d(g("sent")), recvd=_d(g("recvd")), priority=g("priority"),
                         group=str(g("group") or "").strip(), region=str(g("region") or "").strip()))
    return rows


def check_rows(rows, acked=frozenset(), today=None):
    today = today or dt.date.today()
    issues = []

    def add(r, code, level, msg, options=()):
        if (r["no"], code) in acked:
            return
        issues.append(dict(code=code, no=r["no"], name=r["name"], row=r["row"], level=level, message=msg,
                           options=list(options)))

    seen = {}
    for r in rows:
        seen.setdefault(r["no"], []).append(r)
        erp, th, diff, st = r["erp"], r["theirs"], r["diff"], r["status"]
        got_reply = bool(st) or r["bc_status"] == "received" or th is not None
        if th is not None and erp is not None:
            exp = round(erp - th, 2)
            if diff is None or abs(exp - diff) > TOL:
                add(r, "arith", "fix", f"Difference is {diff if diff is not None else 'empty'} but ERP minus the customer's balance is {exp}.",
                    [("recalc", f"Use {exp}")])
            if st == "clear" and abs(exp) > TOL:
                add(r, "clear_diff", "look", f"Marked Clear but the balances differ by {exp:,.2f}. Any difference is a dispute.",
                    [("dispute", "Make it Dispute"), ("accept", "It is correct, keep Clear")])
            if st == "dispute" and abs(exp) <= TOL:
                add(r, "dispute_same", "look", "Marked Dispute but the balances are equal.",
                    [("clear", "Make it Clear"), ("accept", "It is correct, keep Dispute")])
            if erp and abs(exp) > abs(erp) * BIG_SHARE and abs(exp) > 1000:
                add(r, "big", "look", f"The difference ({exp:,.2f}) is more than half the balance. Check it is not a period total or a typing slip.",
                    [("accept", "Checked, it is right")])
            if erp and th and (erp > 0) != (th > 0):
                add(r, "sign", "look", "The customer's balance has the opposite sign to ours.", [("accept", "Checked, it is right")])
        if st in ("clear", "dispute") and r["bc_status"] != "received":
            add(r, "no_received", "look", f"Marked {st.title()} but Balance Confirmation Status is not 'Received'.")
        if r["bc_status"] == "received" and not st:
            add(r, "received_no_status", "look", "Status says Received but Clear / Dispute is empty.")
        if st == "dispute" and th is None:
            add(r, "dispute_no_bal", "look", "Dispute without the customer's balance.")
        if r["recvd"] and r["sent"] and r["recvd"] < r["sent"]:
            add(r, "date_order", "look", f"Received ({r['recvd']:%d-%m-%Y}) is before sent ({r['sent']:%d-%m-%Y}).")
        if r["recvd"] and r["recvd"] > today:
            add(r, "date_future", "look", f"Received date {r['recvd']:%d-%m-%Y} is in the future.")
    for no, rs in seen.items():
        if len(rs) > 1:
            add(rs[1], "dup", "look", f"Customer {no} appears {len(rs)} times (rows {', '.join(str(x['row']) for x in rs)}).")
    return issues


def mis_tieout(path, division_map):
    """MIS cells that are typed numbers vs counts worked out from the list. Formula cells are skipped."""
    wb = openpyxl.load_workbook(path)
    ws = workbook._sheet(wb, "consolidated")
    hrow, cols = workbook._header(ws)
    rows = []
    for r in range(hrow + 1, ws.max_row + 1):
        if ws.cell(r, cols["no"]).value in (None, ""):
            continue
        g = lambda k: ws.cell(r, cols[k]).value if k in cols else None
        rows.append(dict(region=str(g("region") or ""), group=str(g("group") or "").strip(), priority=g("priority"),
                         received=bool(g("recvd")) or str(g("bc_status") or "").lower() == "received"))
    counts, _ = workbook.mis_counts(rows, division_map)
    mis = workbook._sheet(wb, "mis")
    diffs, region = [], None
    for r in range(4, mis.max_row + 1):
        a, b = mis.cell(r, 1).value, mis.cell(r, 2).value
        if a and "total" in str(a).lower():
            region = None
            continue
        if a:
            region = str(a).strip().upper()
        if not (region and b):
            continue
        t = counts.get((region, workbook.ALIASES.get(str(b).strip().upper(), str(b).strip().upper())),
                       dict(total=0, p1=0, p3=0, received=0))
        for col, key, label in ((3, "total", "Total"), (4, "p1", "Priority 1"), (6, "p3", "Priority 3"), (7, "received", "Received")):
            v = mis.cell(r, col).value
            if isinstance(v, (int, float)) and v != t[key]:
                diffs.append(dict(region=region.title(), division=str(b), column=label, in_mis=v, from_list=t[key]))
    return diffs
