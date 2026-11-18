"""Half-yearly ERP dump (Business Central 'Customer Wise Details' + 'Customer - Aging Detailed')
-> a fresh consolidated list in his own template.

Mechanical columns are filled from the dump. Columns that exist only in his head or in last period's list
(Email id, Priority, Regional Team full name, SENT ON, ...) are carried over from the previous list by customer No.
or left blank and reported. Nothing is guessed.
"""
import datetime as dt
import shutil
from pathlib import Path

import openpyxl

from . import workbook

# Raw Customer Posting Group in the dump -> label used in his list and MIS (checked against his MIS counts).
RAW_GROUPS = {"AUTOMOTIVE": "AUTOMOTIVE (IPOL)", "REPSOL": "AUTOMOTIVE (REPSOL)", "INDUSTRIAL": "INDUSTRIAL (CORE)",
              "RPO": "RPO", "BITUMEN": "BITUMEN", "BOILS": "Base Oils", "FO": "Fuel Oil", "RPO TRADING": "RPO TRADING"}
REGIONS = ("EAST", "NORTH", "SOUTH", "WEST")
ZONE_LETTER = {"E": "EAST", "N": "NORTH", "S": "SOUTH", "W": "WEST"}
OUT_HEADERS = ["Sr No", "Customer Posting Group", "Regional Team", "Region", "No.", "Name",
               "BALANCE AS PER ERP as on {asof}", "BALANCE AS PER Customer", "DIFFERENCE", "Clear / Dispute", "Remarks",
               "SENT ON", "RECVD ON", "Email Send/Pending Status", "Balance Confirmation to send Yes/No", "Priority",
               "Sales Person Name", "Email id", "Balance Confirmation Status", "Sales Person Region Code"]


def region_from_code(code):
    c = str(code or "").strip().upper()
    for r in REGIONS:
        if c.startswith(r):
            return r
    if c.startswith("ZONE") and len(c) >= 6 and c[5] in ZONE_LETTER:
        return ZONE_LETTER[c[5]]
    return ""


def read_dump(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet = next((n for n in wb.sheetnames if n.strip().lower().startswith("customer wise")), None)
    if not sheet:
        raise ValueError("This does not look like the ERP dump: no sheet called 'Customer Wise Details'.")
    rows = list(wb[sheet].iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if r and "no." in [str(c).strip().lower() for c in r if c])
    head = [str(c).strip().lower() if c else "" for c in rows[hi]]
    col = lambda *names: next((head.index(n) for n in names if n in head), None)
    ix = dict(no=col("no."), name=col("name"), excl=col("exclude customer yes/no"), reason=col("reason (if yes)"),
              team=col("regional account name"), sp=col("salesperson name"), spcode=col("salesperson region code"),
              group=col("customer posting group"), bal=col("remaining amt. (lcy)"))
    if None in (ix["no"], ix["name"], ix["group"], ix["bal"]):
        raise ValueError("The dump is missing a column I need (No., Name, Customer Posting Group, Remaining Amt.).")
    out = []
    for r in rows[hi + 1:]:
        if r and r[ix["no"]]:
            out.append({k: (r[v] if v is not None else None) for k, v in ix.items()})
    # Region = sales region code (WEST 1 -> WEST; matches his MIS to within 11 of 1392). Fallback: Region on the aging sheet.
    region_of, asof = {}, None
    aging = next((n for n in wb.sheetnames if "aging" in n.lower()), None)
    if aging:
        arows = wb[aging].iter_rows(values_only=True)
        ahead = None
        for r in arows:
            if ahead is None:
                if r and r[0] and str(r[0]).strip().lower().startswith("as of"):
                    asof = str(r[0]).split(":", 1)[-1].strip()
                if r and r[0] and str(r[0]).strip().lower() == "no.":
                    ahead = [str(c).strip().lower() if c else "" for c in r]
                    ri = ahead.index("region") if "region" in ahead else None
                continue
            if r and r[0] and ri is not None and r[ri] and r[0] not in region_of:
                region_of[r[0]] = str(r[ri]).strip().upper()
    return out, region_of, asof


def _fmt_date(v):
    """'30-09-26' or '30-09-2026' -> '30-09-2026'."""
    import re
    m = re.match(r"^\s*(\d{1,2})[-./](\d{1,2})[-./](\d{2,4})\s*$", str(v or ""))
    if not m:
        return str(v or "").strip()
    d, mo, y = m.groups()
    return f"{int(d):02d}-{int(mo):02d}-{int(y) + 2000 if len(y) == 2 else int(y)}"


def read_previous(path):
    """No. -> what the previous list knew that the dump does not."""
    if not path or not Path(path).exists():
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = workbook._sheet(wb, "consolidated")
    hrow, _ = workbook._header(ws)
    head = {str(c.value).strip().lower(): c.column for c in ws[hrow] if c.value}
    get = lambda r, n: ws.cell(r, head[n]).value if n in head else None
    out = {}
    for r in range(hrow + 1, ws.max_row + 1):
        no = get(r, "no.")
        if no:
            out[str(no).strip()] = dict(team=get(r, "regional team"), email=get(r, "email id"), priority=get(r, "priority"))
    return out


def build(dump_path, template_path, out_path, division_map=None, previous_path=None, regional_teams=None, asof_label=None):
    """Make the new consolidated workbook. Returns a report dict for the screen."""
    rows, region_of, asof = read_dump(dump_path)
    prev = read_previous(previous_path or template_path)
    teams = {k.strip().lower(): v for k, v in (regional_teams or {}).items()}
    wb = openpyxl.load_workbook(template_path)
    ws = workbook._sheet(wb, "consolidated")
    hrow, cols = workbook._header(ws)
    # keep the header row and its styles; replace everything below it
    style_row = hrow + 1
    styles = {c: ws.cell(style_row, c)._style for c in range(1, ws.max_column + 1)} if ws.max_row > hrow else {}
    ws.delete_rows(hrow + 1, ws.max_row)
    hdr = {str(c.value).strip().lower(): c.column for c in ws[hrow] if c.value}
    erp_col = next(c for t, c in hdr.items() if t.startswith("balance as per erp"))
    label_date = _fmt_date(asof_label or asof)
    if label_date:
        ws.cell(hrow, erp_col, f"BALANCE AS PER ERP as on {label_date}")
    report = dict(written=0, dropped=[], new_customers=[], no_priority=0, no_email=0, unknown_team=set(),
                  carried=0, unmapped_region=[], excluded_flag=0)
    gm = {k.upper(): v for k, v in (division_map or RAW_GROUPS).items()}
    n = 0
    for d in rows:
        group = str(d["group"] or "").strip()
        label = gm.get(group.upper())
        if label is None:
            report["dropped"].append((d["no"], d["name"], group))
            continue
        n += 1
        no = str(d["no"]).strip()
        old = prev.get(no)
        team_raw = str(d["team"] or "").strip()
        team = teams.get(team_raw.lower()) or (old or {}).get("team") or team_raw
        if team_raw and not (teams.get(team_raw.lower()) or (old or {}).get("team")):
            report["unknown_team"].add(team_raw)
        region = region_from_code(d["spcode"]) or region_of.get(d["no"], "")
        if not region:
            report["unmapped_region"].append(no)
        yes = str(d["excl"] or "").strip().lower() == "yes"
        report["excluded_flag"] += yes
        vals = {"sr no": n, "customer posting group": label, "regional team": team, "region": region, "no.": no,
                "name": d["name"], erp_col: d["bal"], "balance confirmation to send yes/no": "No" if yes else "Yes",
                "priority": (old or {}).get("priority"), "sales person name": d["sp"],
                "email id": (old or {}).get("email"), "sales person region code": d["spcode"]}
        if yes and d.get("reason"):
            vals["remarks"] = d["reason"]
        for key, v in vals.items():
            c = key if isinstance(key, int) else next((cc for t, cc in hdr.items() if t == key), None)
            if c:
                ws.cell(hrow + n, c, v)
        for c, st in styles.items():
            ws.cell(hrow + n, c)._style = st
        if old is None:
            report["new_customers"].append((no, d["name"]))
        else:
            report["carried"] += 1
        if not (old or {}).get("priority"):
            report["no_priority"] += 1
        if not (old or {}).get("email"):
            report["no_email"] += 1
    report["written"] = n
    report["unknown_team"] = sorted(report["unknown_team"])
    # MIS starts again from zero received (counts, not his formulas).
    unmapped = workbook.fill_mis(wb, ws, hrow, dict(no=hdr["no."], region=hdr["region"], group=hdr["customer posting group"],
                                                    priority=hdr.get("priority", 16), recvd=hdr.get("recvd on", 13),
                                                    bc_status=hdr.get("balance confirmation status", 19)),
                                 {**{k: v for k, v in RAW_GROUPS.items()}, **(division_map or {}),
                                  **{v.upper(): v for v in RAW_GROUPS.values()}})
    report["mis_unmapped"] = sorted(unmapped)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = out_path.with_suffix(".tmp.xlsx")
    wb.calculation.fullCalcOnLoad = True  # Excel recalculates his formulas when opened
    wb.save(tmp)
    shutil.move(tmp, out_path)
    report["path"] = str(out_path)
    return report
