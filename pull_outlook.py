"""Copy new mails from a classic-Outlook folder into the watched folder, with no browser.
Run it by hand, or from Windows Task Scheduler every 15 minutes:   python pull_outlook.py
Uses config.json (or BC_CONFIG): outlook_folder, outlook_save_to (or the first source). Only reads Outlook."""
import json
import os
from pathlib import Path

from bc import outlook

HERE = Path(__file__).resolve().parent
cfg = json.loads(Path(os.environ.get("BC_CONFIG", HERE / "config.json")).read_text())
data = Path(cfg["data_folder"]); data.mkdir(parents=True, exist_ok=True)
state = data / "outlook_seen.json"
seen = set(json.loads(state.read_text())) if state.exists() else set()
src = cfg.get("sources") or [{"path": cfg["shared_folder"]}]
dest = Path(cfg.get("outlook_save_to") or src[0]["path"])
def _known():
    from bc import workbook
    from bc.rules import EMAIL_RE
    emails, domains, numbers = set(), set(), set()
    for no, c in workbook.load_customers(cfg["consolidated_list"]).items():
        numbers.add(str(no).upper())
        for a in EMAIL_RE.findall(str(c.get("email") or "").lower()):
            emails.add(a)
            domains.add(a.split("@")[-1])
    return dict(emails=emails, domains=domains, numbers=numbers)


saved, seen, rep = outlook.pull(cfg.get("outlook_folders") or cfg.get("outlook_folder", "Inbox"), dest, seen, keywords=cfg.get("outlook_keywords", []), known=_known(), maybe_dir=Path(cfg.get("outlook_maybe_folder") or data / "maybe"))
state.write_text(json.dumps(sorted(seen)))
print(f"Outlook folder: {rep['in_folder']} item(s) found, {rep['saved']} new saved to {dest}, {rep['already']} already pulled, "
      f"{rep['not_mail']} not ordinary mail, {rep['filtered']} left out, {rep['maybe']} put in the maybe list, {len(rep['failed'])} failed")
print("  why kept:", rep["why"])
for f in rep["failed"]:
    print("  FAILED:", f)
