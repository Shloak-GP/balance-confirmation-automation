"""Run the tool on the REAL files in ./real_data (the owner's folder):   python3 run_real.py
   python3 run_real.py --fresh   starts again from nothing (forgets everything you saved).

real_data/replies/                          the 19 real reply mails / PDFs (read-only: never changed)
real_data/Consolidated_List_from_ERP_dump.xlsx  the list built from the ERP dump (never changed)
real_work/                                  the tool's own copy of the list, backups, reports and database
"""
import json
import os
import shutil
import sys
import threading
import webbrowser
from pathlib import Path

HERE = Path(__file__).resolve().parent
real, work = HERE / "real_data", HERE / "real_work"
if "--fresh" in sys.argv and work.exists():
    shutil.rmtree(work)
work.mkdir(exist_ok=True)
live = work / "Consolidated_List.xlsx"
if not live.exists():
    shutil.copy(real / "Consolidated_List_from_ERP_dump.xlsx", live)
cfg = json.loads((HERE / "config.example.json").read_text())
cfg.update(shared_folder=str(real / "replies"), sources=[{"name": "Shared folder", "path": str(real / "replies")}],
           consolidated_list=str(live), output_folder=str(work / "reports"), data_folder=str(work / "data"),
           report_link_prefix="", write_mode="in_place", host="127.0.0.1", port=8080, period="Sept 2026")
ov = HERE / "local_overrides.json"          # optional: e.g. {"outlook_folder": "Inbox/Balance Confirmation", "thread_subject": "..."}
if ov.exists():
    cfg.update(json.loads(ov.read_text()))
(work / "config.json").write_text(json.dumps(cfg, indent=1))
os.environ["BC_CONFIG"] = str(work / "config.json")
os.chdir(HERE)
sys.path.insert(0, str(HERE))
import web  # noqa: E402

print("\nReal replies:", real / "replies")
print("Tool's working list:", live)
print("Open http://127.0.0.1:8080  (Ctrl+C to stop)\n")
threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:8080")).start()
web.app.run(host="127.0.0.1", port=8080)
