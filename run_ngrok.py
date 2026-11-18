"""Start the tool on the real-data pack AND an ngrok link, with a password. One command:  python run_ngrok.py
Needs ngrok installed and signed in once (ngrok config add-authtoken ...). Ctrl+C stops both; the link dies."""
import json
import secrets
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
ov = HERE / "local_overrides.json"
cfg = json.loads(ov.read_text()) if ov.exists() else {}
if not cfg.get("access_password"):
    cfg["access_password"] = secrets.token_urlsafe(9)
    ov.write_text(json.dumps(cfg, indent=1))
pw = cfg["access_password"]

ngrok = shutil.which("ngrok")
if not ngrok:
    sys.exit("ngrok was not found. Install it from ngrok.com, run:  ngrok config add-authtoken YOUR_TOKEN  then try again.")

app = subprocess.Popen([sys.executable, str(HERE / "run_real.py")] + (["--fresh"] if "--fresh" in sys.argv else []), cwd=HERE)
tun = subprocess.Popen([ngrok, "http", "8080"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
url = None
for _ in range(40):
    time.sleep(1)
    try:
        t = json.load(urllib.request.urlopen("http://127.0.0.1:4040/api/tunnels", timeout=2))["tunnels"]
        url = next((x["public_url"] for x in t if x["public_url"].startswith("https")), None)
        if url:
            break
    except Exception:
        pass
print("\n" + "=" * 60)
if url:
    print("Send him this link:  ", url)
    print("Password (send it separately):", pw)
    print("He may see an ngrok 'Visit site' page first: he clicks it.")
else:
    print("ngrok started but no link was found. Check that ngrok is signed in (ngrok config add-authtoken ...).")
print("Keep this window open. Ctrl+C stops everything and the link dies.")
print("=" * 60 + "\n")
try:
    app.wait()
except KeyboardInterrupt:
    pass
finally:
    for p in (tun, app):
        try:
            p.terminate()
        except Exception:
            pass
