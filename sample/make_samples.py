"""Make FAKE test data: a filled consolidated list in his template, and a shared folder of replies.

All names, codes and amounts are invented. Run:  python sample/make_samples.py <out_dir>
"""
import random
import shutil
import sys
from email.message import EmailMessage
from pathlib import Path

import openpyxl
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "his_template.xlsx"  # his empty MIS + consolidated template (headers only)

GROUPS = {"NORTH": ["AUTO-IPOL", "AUTO-REPSOL", "IND-CORE", "RPO", "BITUMEN", "BOILS", "FUEL OIL"],
          "WEST": ["AUTO-IPOL", "AUTO-REPSOL", "IND-CORE", "RPO", "BITUMEN", "BOILS", "RPO TRADING"]}
FIRST = ["Shree Ganesh", "Patel", "Kaveri", "Sai Krupa", "Om", "Jai Bharat", "Laxmi", "Royal", "Mahalaxmi",
         "Sunrise", "Narmada", "Ganga", "Annapurna", "Balaji", "Hind", "Vijay", "Krishna", "Apex", "Unity", "Star"]
LAST = ["Traders", "Auto Spares", "Lubricants", "Enterprises", "Motors", "Agencies", "Industries", "Oil Co"]


def make_list(out):
    random.seed(7)
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Consolidated List "]
    r, n = 5, 0
    custs = []
    for region, groups in GROUPS.items():
        for i in range(20):
            n += 1
            no = f"C{region[0]}{10000 + n}"
            name = f"{FIRST[(n * 7) % len(FIRST)]} {LAST[(n * 3) % len(LAST)]}" + ("" if n < 21 else f" {region.title()}")
            erp = round(random.uniform(20000, 900000), 2) if i % 9 != 8 else 0.0
            prio = 3 if erp == 0 else random.choice([1, 2, 2])
            vals = [n, groups[i % len(groups)], f"{region.title()} Team", region, no, name, erp, None, None, None, None,
                    "2026-10-05", None, "Sent", "Yes", prio, "Sales Person", "accounts@example.com", "Pending", region[:2], None]
            for c, v in enumerate(vals, 1):
                ws.cell(r, c, v)
            custs.append(dict(no=no, name=name, erp=erp))
            r += 1
    wb.save(out)
    return custs


def form_pdf(path, cust, line, sign=True):
    c = canvas.Canvas(str(path), pagesize=A4)
    y = 780
    for t in ["GP PETROLEUMS LTD - BALANCE CONFIRMATION", f"Customer No. {cust['no']}   {cust['name']}", "",
              f"As per our books the balance as on 30.09.2026 is Rs {inr(cust['erp'])} Dr.", "", line,
              "", "Stamp & signature of customer" if sign else ""]:
        c.drawString(60, y, t); y -= 24
    c.save()


def scanned_pdf(path, cust, handwritten):
    img = Image.new("RGB", (1240, 1754), "white")
    d = ImageDraw.Draw(img)
    f = ImageFont.load_default(size=28)
    d.text((80, 100), "BALANCE CONFIRMATION", fill="black", font=f)
    d.text((80, 160), f"Customer No. {cust['no']}  {cust['name']}", fill="black", font=f)
    d.text((80, 260), f"As per our books the balance is Rs {inr(cust['erp'])} Dr.", fill="black", font=f)
    d.text((80, 360), "Balance as per your books:", fill="black", font=f)
    d.text((140, 420), handwritten, fill=(20, 60, 160), font=ImageFont.load_default(size=40))
    img.save(path, "PDF", resolution=150)


def eml(path, frm, subject, body, attach=None):
    m = EmailMessage()
    m["From"], m["To"], m["Subject"], m["Date"] = frm, "rm.north@example.com", subject, "Mon, 16 Nov 2026 09:30:00 +0530"
    m.set_content(body)
    if attach:
        m.add_attachment(Path(attach).read_bytes(), maintype="application", subtype="pdf", filename=Path(attach).name)
    Path(path).write_bytes(bytes(m))


def inr(v):
    whole, frac = f"{v:.2f}".split(".")
    head, tail = whole[:-3], whole[-3:]
    g = []
    while len(head) > 2:
        g.insert(0, head[-2:]); head = head[:-2]
    if head:
        g.insert(0, head)
    return ",".join(g + [tail]) + "." + frac if g else tail + "." + frac


def main(out):
    out = Path(out)
    if out.exists():
        shutil.rmtree(out)
    shared = out / "shared_folder"
    (shared / "North").mkdir(parents=True)
    (shared / "West").mkdir(parents=True)
    custs = make_list(out / "Consolidated_List_FAKE.xlsx")
    c = custs
    # 1 dispute, typed amount on form, ID + name in file name
    form_pdf(shared / "North" / f"{c[0]['no']} {c[0]['name']}.pdf", c[0],
             f"[x] We do not agree. Balance as per our books Rs {inr(c[0]['erp'] - 70000)} Dr")
    # 2 same file saved again by another RM -> skipped as duplicate
    shutil.copy(shared / "North" / f"{c[0]['no']} {c[0]['name']}.pdf", shared / "West" / "copy of reply.pdf")
    # 3 "no changes" mail with signed form attached
    tmp = out / "tmp_form.pdf"
    form_pdf(tmp, c[1], "[x] We confirm the above balance")
    eml(shared / "North" / "Re Balance confirmation.eml", "accounts@cust2.example",
        f"Re: Balance confirmation {c[1]['no']} {c[1]['name']}", "Dear Sir,\nNo changes from our side.\nRegards", tmp)
    # 4 scanned, handwritten amount -> owner types
    scanned_pdf(shared / "North" / f"{c[2]['no']} {c[2]['name']} scan.pdf", c[2], f"Rs {inr(c[2]['erp'] - 1520.5)}/-")
    # 5 "confirmed except" -> flagged, no automatic clear
    eml(shared / "West" / "balance conf.eml", "finance@cust.example", f"Balance confirmation {c[22]['no']}",
        f"Dear Sir,\nWe confirm the balance except bill no. 4471 for Rs 12,400.00 which is not in our books.\nRegards\n{c[22]['name']}")
    # 6 no ID anywhere, name only
    form_pdf(shared / "West" / "confirmation.pdf", dict(c[23], no=""),
             f"Balance as per our books Rs {inr(c[23]['erp'])} Dr. We confirm.")
    # 7 ID belongs to another customer than the name on the form
    bad = dict(c[24], no=c[25]["no"])
    form_pdf(shared / "West" / f"{c[25]['no']} reply.pdf", bad, f"Balance as per our books Rs {inr(c[24]['erp'])}")
    tmp.unlink()
    print("customers", len(custs), "->", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else HERE / "fake")
