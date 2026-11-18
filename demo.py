"""Try the tool on FAKE data, no setup:  python demo.py   (then it opens http://127.0.0.1:8080)

Makes a fake shared folder + consolidated list in ./demo_data, points the tool at it and starts it.
Run again to start fresh.
"""
import json
import os
import shutil
import sys
import threading
import webbrowser
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "sample"))
import make_samples  # noqa: E402

demo = HERE / "demo_data"
if demo.exists():
    shutil.rmtree(demo)
make_samples.main(demo / "fake")
# Fake "his mail": one reply that is also in the shared folder (will be merged), one that only he has.
mail = demo / "fake" / "his_mail"
mail.mkdir(parents=True)
src = next((demo / "fake" / "shared_folder").rglob("*.eml"))
shutil.copy(src, mail / "copy of a reply the RM also saved.eml")
(mail / "only in his mail.eml").write_text(
    "From: accounts@example.com\nSubject: Balance confirmation CN10005\n\nDear Sir, we confirm. No changes.\n")
cfg = json.loads((HERE / "config.example.json").read_text())
cfg.update(shared_folder=str(demo / "fake" / "shared_folder"),
           sources=[{"name": "Shared folder", "path": str(demo / "fake" / "shared_folder")},
                    {"name": "His mail", "path": str(mail)}],
           consolidated_list=str(demo / "fake" / "Consolidated_List_FAKE.xlsx"),
           output_folder=str(demo / "reports"), data_folder=str(demo / "data"),
           host="127.0.0.1", port=8080, boss_email="boss@example.com")
(demo / "config.json").write_text(json.dumps(cfg, indent=1))
os.environ["BC_CONFIG"] = str(demo / "config.json")
os.chdir(HERE)
sys.path.insert(0, str(HERE))
import web  # noqa: E402

print("\nFake replies are in:", demo / "fake" / "shared_folder")
print("Excel output goes to:", demo / "reports")
print("Open http://127.0.0.1:8080  (Ctrl+C to stop)\n")
threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:8080")).start()
web.app.run(host="127.0.0.1", port=8080)
