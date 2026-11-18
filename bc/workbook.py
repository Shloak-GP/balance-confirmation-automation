"""Read the owner's consolidated list and write his two reports, in his own template.

The tool never edits his original file. It reads the base list once, keeps every
decision in its own database, and regenerates the output workbook from base + decisions,
so the file and the tool can never drift apart.
"""
import datetime as dt
import os
import re
import shutil
from pathlib import Path

import openpyxl

from . import rules

# Header text in his sheet (matched by "starts with", case-insensitive).
H = {
    "group": "customer posting group", "team": "regional team", "region": "region", "no": "no.",
    "name": "name", "erp": "balance as per erp", "theirs": "balance as per customer",
    "diff": "difference", "status": "clear / dispute", "remarks": "remarks", "sent": "sent on",
    "recvd": "recvd on", "priority": "priority", "email": "email id", "bc_status": "balance confirmation status",
}


def _sheet(wb, prefix):
    for n in wb.sheetnames:
        if n.strip().lower().startswith(prefix):
            return wb[n]
    raise KeyError(f"No sheet starting with '{prefix}' in {wb.sheetnames}")


def _header(ws):
    for row in ws.iter_rows(min_row=1, max_row=15):
        texts = {str(c.value).strip().lower(): c.column for c in row if c.value is not None}
        if "no." in texts:
            cols = {}
            for key, start in H.items():
                for t, col in texts.items():
                    if (t == start if key in ("no", "name", "region") else t.startswith(start)):
                        cols[key] = col
                        break
            return row[0].row, cols
    raise ValueError("Could not find the header row (a cell saying 'No.').")


def load_customers(path):
    """Return {No.: {...}} from the consolidated list."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = _sheet(wb, "consolidated")
    hrow, cols = _header(ws)
    out = {}
    for r in range(hrow + 1, ws.max_row + 1):
        no = ws.cell(r, cols["no"]).value
        if no in (None, ""):
            continue
        no = str(no).strip()
        get = lambda k: ws.cell(r, cols[k]).value if k in cols else None
        out[no] = dict(row=r, no=no, name=str(get("name") or "").strip(), region=str(get("region") or "").strip(),
                       group=str(get("group") or "").strip(), priority=get("priority"), erp=_num(get("erp")),
                       recvd=get("recvd"), status=get("status"), theirs=_num(get("theirs")),
                       email=str(get("email") or ""))
    return out


def _num(v):
    if v is None or v == "":
        return None
    try:
        return round(float(str(v).replace(",", "")), 2)
    except ValueError:
        return None


class FileLocked(Exception):
    """His Excel is open in Excel right now; try again later."""


def _apply(wb, decisions, division_map, fill_counts=False, received_formula=False):
    ws = _sheet(wb, "consolidated")
    hrow, cols = _header(ws)
    rows = {}
    for r in range(hrow + 1, ws.max_row + 1):
        v = ws.cell(r, cols["no"]).value
        if v not in (None, ""):
            rows[str(v).strip()] = r
    missing = []
    for d in decisions:
        r = rows.get(d["no"])
        if r is None:
            missing.append(d["no"])
            continue
        erp = _num(ws.cell(r, cols["erp"]).value)
        theirs = d["theirs"]
        put = lambda k, v: ws.cell(r, cols[k], v) if k in cols else None
        # Only the tool's own columns. Every other column is his and is never touched.
        put("theirs", theirs)
        put("diff", rules.difference(erp, theirs))
        put("status", "Clear" if d["status"] == "CLEAR" else "Dispute")
        if d.get("remarks"):
            put("remarks", d["remarks"])
        put("recvd", dt.datetime.fromisoformat(d["saved_at"]).date() if d.get("received_on") is None
            else dt.date.fromisoformat(d["received_on"]))
        if "recvd" in cols:
            ws.cell(r, cols["recvd"]).number_format = "DD-MM-YYYY"
        put("bc_status", "Received")
    if fill_counts:
        fill_mis(wb, ws, hrow, cols, division_map)
    elif received_formula:
        add_mis_formulas(wb, division_map)
    return missing


def write_outputs(base_path, decisions, out_dir, division_map, fill_counts=False, received_formula=False):
    """Copy mode: a fresh dated workbook in out_dir made from base + decisions. His file is never touched."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(base_path)  # keeps his formatting and formulas
    missing = _apply(wb, decisions, division_map, fill_counts, received_formula)
    path = out_dir / f"Balance_Confirmation_{dt.date.today():%Y-%m-%d}.xlsx"
    tmp = path.with_suffix(".tmp.xlsx")
    wb.calculation.fullCalcOnLoad = True  # Excel recalculates his formulas when opened
    wb.save(tmp)
    shutil.move(tmp, path)  # never leave a half-written file behind
    return path, missing


def write_mis_only(source_xlsx, out_dir, division_map=None):
    """The file that goes to the boss: ONLY the MIS sheet (no customer-level list, no balances).

    The MIS cells are plain numbers (his % formulas stay), so the consolidated sheet can be dropped safely.
    Always written as a new dated copy; his files are never touched.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(source_xlsx)
    if division_map:
        freeze_mis_links(wb, division_map)
    for name in list(wb.sheetnames):
        if not name.strip().lower().startswith("mis"):
            del wb[name]
    path = out_dir / f"MIS_Balance_Confirmation_{dt.date.today():%Y-%m-%d}.xlsx"
    tmp = path.with_suffix(".tmp.xlsx")
    wb.calculation.fullCalcOnLoad = True  # Excel recalculates his formulas when opened
    wb.save(tmp)
    shutil.move(tmp, path)
    return path


def freeze_mis_links(wb, division_map):
    """If his MIS counts are formulas that read the Consolidated sheet (COUNTIFS etc.), dropping that sheet for the
    boss's copy would turn them into #REF!. Replace exactly those cells by the same numbers, counted by the tool.
    Same-sheet formulas (P2, Pending, %, region totals) are kept live."""
    ws = _sheet(wb, "mis")
    cons = _sheet(wb, "consolidated")
    hrow, cols = _header(cons)
    rows = []
    for r in range(hrow + 1, cons.max_row + 1):
        if cons.cell(r, cols["no"]).value in (None, ""):
            continue
        g = lambda k: cons.cell(r, cols[k]).value if k in cols else None
        rows.append(dict(region=str(g("region") or ""), group=str(g("group") or "").strip(), priority=g("priority"),
                         received=bool(g("recvd")) or str(g("bc_status") or "").lower() == "received"))
    counts, _ = mis_counts(rows, division_map)
    region = None
    for r in range(4, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a and "total" in str(a).lower():
            region = None
            continue
        if a:
            region = str(a).strip().upper()
        if not (region and b):
            continue
        t = counts.get((region, ALIASES.get(str(b).strip().upper(), str(b).strip().upper())),
                       dict(total=0, p1=0, p2=0, p3=0, received=0))
        vals = [t["total"], t["p1"], t["p2"], t["p3"], t["received"], t["total"] - t["p3"] - t["received"]]
        for col, v in zip(range(3, 9), vals):
            f = ws.cell(r, col).value
            if isinstance(f, str) and f.startswith("=") and "!" in f:
                ws.cell(r, col, v)


def is_open_in_excel(path):
    """Excel puts a hidden '~$name.xlsx' next to a file while it is open (also on shared drives)."""
    p = Path(path)
    return (p.parent / ("~$" + p.name[2:] if len(p.name) > 2 else "~$" + p.name)).exists() or \
           (p.parent / ("~$" + p.name)).exists()


def write_in_place(live_path, decisions, backup_dir, division_map, keep=30, fill_counts=False, received_formula=False):
    """In-place mode: update his real file, only the tool's columns + the MIS counts.

    He can edit every other column in Excel. If the file is open, raises FileLocked (nothing written).
    A timestamped backup is taken before every write.
    """
    live = Path(live_path)
    if is_open_in_excel(live):
        raise FileLocked(str(live))
    backup_dir = Path(backup_dir)
    backup_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(live, backup_dir / f"{live.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}{live.suffix}")
    for old in sorted(backup_dir.glob(f"{live.stem}_*{live.suffix}"))[:-keep]:
        old.unlink()
    wb = openpyxl.load_workbook(live)
    missing = _apply(wb, decisions, division_map, fill_counts, received_formula)
    tmp = live.with_name(live.stem + ".balcheck-tmp" + live.suffix)
    wb.calculation.fullCalcOnLoad = True  # Excel recalculates his formulas when opened
    wb.save(tmp)
    try:
        os.replace(tmp, live)  # atomic; fails on Windows if Excel holds the file
    except PermissionError as e:
        tmp.unlink(missing_ok=True)
        raise FileLocked(str(live)) from e
    return live, missing


# His MIS uses two spellings for the same division in different regions.
ALIASES = {"BOILS": "BASE OILS"}


def _prio(v):
    """'Priority 1 (Delinquent Customers)' -> '1'; 'P2' -> '2'; blank -> None."""
    m = re.search(r"(?i)\bp(?:riority)?\s*([123])\b", str(v or ""))
    return m.group(1) if m else None


def add_mis_formulas(wb, division_map, which=("total", "p1", "p3", "received")):
    """Put live formulas in the MIS count cells (C Total, D P1, F P3, G Received), counting rows of the consolidated
    list per region and division. P2, Pending and % stay his own formulas. A cell that already holds a formula is
    left alone. Groups and priority labels are listed one by one (no wildcards), so nothing is counted twice."""
    from openpyxl.utils import get_column_letter as L
    cons = _sheet(wb, "consolidated")
    hrow, cols = _header(cons)
    need = {"total": ("region", "group"), "p1": ("region", "group", "priority"),
            "p3": ("region", "group", "priority"), "received": ("region", "group", "recvd")}
    which = [w for w in which if all(k in cols for k in need[w])]
    groups, prios = {}, {"1": set(), "3": set()}
    for r in range(hrow + 1, cons.max_row + 1):
        if cons.cell(r, cols["no"]).value in (None, ""):
            continue
        g = str(cons.cell(r, cols["group"]).value or "").strip()
        div = division_map.get(g.upper(), division_map.get(g))
        if "priority" in cols:
            raw = cons.cell(r, cols["priority"]).value
            p = _prio(raw)
            if p in prios and str(raw).strip():
                prios[p].add(str(raw).strip())
        if div is None:
            continue
        d = ALIASES.get(div.strip().upper(), div.strip().upper())
        groups.setdefault(d, set()).add(g)
    q = "'" + cons.title.replace("'", "''") + "'!"
    last = max(cons.max_row + 2000, 5000)          # room to grow
    rng = lambda k: f"{q}${L(cols[k])}${hrow + 1}:${L(cols[k])}${last}"
    esc = lambda t: str(t).replace('"', '""')
    ws = _sheet(wb, "mis")
    column = {"total": 3, "p1": 4, "p3": 6, "received": 7}
    region, done = None, 0
    for r in range(4, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a and "total" in str(a).lower():
            region = None
            continue
        if a:
            region = str(a).strip().upper()
        if not (region and b):
            continue
        gs = sorted(groups.get(ALIASES.get(str(b).strip().upper(), str(b).strip().upper()), ()))
        for w in which:
            cell = ws.cell(r, column[w])
            if str(cell.value or "").startswith("="):
                continue
            parts = []
            for g in gs:
                base = f'{rng("region")},"{esc(region)}",{rng("group")},"{esc(g)}"'
                if w == "total":
                    parts.append(f"COUNTIFS({base})")
                elif w == "received":
                    parts.append(f'COUNTIFS({base},{rng("recvd")},"<>")')
                else:
                    for pv in sorted(prios[w[1]]):
                        parts.append(f'COUNTIFS({base},{rng("priority")},"{esc(pv)}")')
            cell.value = "=" + "+".join(parts) if parts else 0
            done += 1
    return done


def mis_counts(customers_rows, division_map):
    """customers_rows: iterable of dicts with region, group, priority, received(bool)."""
    counts = {}
    unmapped = set()
    for c in customers_rows:
        div = division_map.get(c["group"].upper(), division_map.get(c["group"], None))
        if div is None:
            unmapped.add(c["group"])
            continue
        d = div.strip().upper()
        key = (c["region"].strip().upper(), ALIASES.get(d, d))
        t = counts.setdefault(key, dict(total=0, p1=0, p2=0, p3=0, received=0))
        t["total"] += 1
        p = _prio(c["priority"])
        if p:
            t["p" + p] += 1
        if c["received"]:
            t["received"] += 1
    return counts, unmapped


def fill_mis(wb, cons_ws, hrow, cols, division_map):
    rows = []
    for r in range(hrow + 1, cons_ws.max_row + 1):
        if cons_ws.cell(r, cols["no"]).value in (None, ""):
            continue
        g = lambda k: cons_ws.cell(r, cols[k]).value if k in cols else None
        rows.append(dict(region=str(g("region") or ""), group=str(g("group") or "").strip(),
                         priority=g("priority"), received=bool(g("recvd")) or str(g("bc_status") or "").lower() == "received"))
    counts, unmapped = mis_counts(rows, division_map)
    ws = _sheet(wb, "mis")
    region = None
    for r in range(4, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if a and "total" in str(a).lower():
            region = None
            continue
        if a:
            region = str(a).strip().upper()
        if not (region and b):
            continue
        label = ALIASES.get(str(b).strip().upper(), str(b).strip().upper())
        t = counts.get((region, label), dict(total=0, p1=0, p2=0, p3=0, received=0))
        # C..H = Total, P1, P2, P3, Received, Pending. His sheet computes P2 (= Total-P1-P3), Pending and
        # % Achievement with formulas: a formula cell is never overwritten, only plain-number cells.
        for col, v in zip(range(3, 9), [t["total"], t["p1"], t["p2"], t["p3"], t["received"],
                                         t["total"] - t["p3"] - t["received"]]):
            if not str(ws.cell(r, col).value or "").startswith("="):
                ws.cell(r, col, v)
    return unmapped
