"""One-time setup for the owner: asks where his files are, writes config.json, makes a desktop start file.
Run by setup.bat (Windows) or:  python setup_wizard.py   (add  --answers answers.json  to skip the questions)."""
import json
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG_OUT = Path(os.environ.get("BC_CONFIG") or HERE / "config.json")


def ask_file(title):
    try:
        import tkinter as tk
        from tkinter import filedialog
        r = tk.Tk(); r.withdraw(); r.attributes("-topmost", True)
        p = filedialog.askopenfilename(title=title, filetypes=[("Excel", "*.xlsx *.xlsm")])
        r.destroy()
        if p:
            return p
    except Exception:
        pass
    return input(f"{title}\n  type or paste the full path: ").strip().strip('"')


def ask_folder(title):
    try:
        import tkinter as tk
        from tkinter import filedialog
        r = tk.Tk(); r.withdraw(); r.attributes("-topmost", True)
        p = filedialog.askdirectory(title=title)
        r.destroy()
        if p:
            return p
    except Exception:
        pass
    return input(f"{title}\n  type or paste the full folder path: ").strip().strip('"')


def main():
    answers = None
    if "--answers" in sys.argv:
        answers = json.loads(Path(sys.argv[sys.argv.index("--answers") + 1]).read_text())
    A = lambda k, fn, *a: (answers[k] if answers and k in answers else fn(*a))
    print("\n=== Balance check: one-time setup ===\n")
    print("1) Pick your consolidated Excel list.")
    lst = A("list", ask_file, "Pick your consolidated list (Excel)")
    print("2) Pick the folder where the reply mails and PDFs are saved.")
    rep = A("replies", ask_folder, "Pick the folder with the reply mails / PDFs")
    docs = Path.home() / "Documents" / "BalanceCheck"
    data, out = docs / "data", docs / "reports"
    boss = A("boss_email", lambda: input("3) Boss's e-mail address (for the MIS mail): ").strip())
    subj = A("thread_subject", lambda: input("4) Subject of the daily mail thread to reply on (press Enter to skip): ").strip())
    outlook = A("outlook_folder", lambda: input("5) Outlook folder with customer replies, e.g. Inbox/Balance Confirmation (Enter to skip): ").strip())
    mode = (answers or {}).get("write_mode", "copy")
    cfg = json.loads((HERE / "config.example.json").read_text())
    cfg.update(shared_folder=rep, sources=[{"name": "Shared folder", "path": rep}], consolidated_list=lst,
               output_folder=str(out), data_folder=str(data), write_mode=mode, host="127.0.0.1", port=8080,
               boss_email=boss, thread_subject=subj, outlook_folder=outlook, report_link_prefix="",
               access_password="", auto_scan_minutes=0)
    cfg.pop("_comment", None)
    CONFIG_OUT.write_text(json.dumps(cfg, indent=1))
    # check that his files can be read, in plain words
    sys.path.insert(0, str(HERE))
    ok = True
    try:
        from bc import workbook
        n = len(workbook.load_customers(lst))
        import openpyxl
        wb = openpyxl.load_workbook(lst)
        workbook._sheet(wb, "mis")
        print(f"\n   OK: read {n} customers from your list and found the MIS sheet.")
    except Exception as e:
        ok = False
        print(f"\n   PROBLEM with the Excel list: {e}\n   Send this message to the person who set this up.")
    files = list(Path(rep).rglob("*")) if Path(rep).exists() else []
    print(f"   OK: {len([f for f in files if f.is_file()])} file(s) in the reply folder." if Path(rep).exists()
          else "   PROBLEM: the reply folder was not found.")
    # desktop start file
    desktop = Path.home() / "Desktop"
    if desktop.exists() and os.name == "nt" and not getattr(sys, "frozen", False):
        (desktop / "Balance Check.bat").write_text(f'@echo off\ncd /d "{HERE}"\ncall start.bat\n')
        print("\n   A 'Balance Check' start file is now on your Desktop.")
    print("\nMode:", "TRIAL (your original Excel is never changed; a dated copy is made in " + str(out) + ")"
          if mode == "copy" else "LIVE (your Excel is updated in place, a backup is kept before every change)")
    print("\nSetup finished." if ok else "\nSetup finished with a problem (see above).")


if __name__ == "__main__":
    main()
