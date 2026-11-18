"""Entry point for the no-install version (built with PyInstaller by build_exe.bat).
First run: asks where his files are (setup wizard). Every run: starts the tool and opens the browser."""
import os
import sys
import threading
import webbrowser
from pathlib import Path

BASE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
os.environ.setdefault("BC_CONFIG", str(BASE / "config.json"))
if getattr(sys, "frozen", False):
    sys.path.insert(0, getattr(sys, "_MEIPASS", str(BASE)))
os.chdir(BASE)

if not Path(os.environ["BC_CONFIG"]).exists():
    import setup_wizard
    setup_wizard.main()
    if not Path(os.environ["BC_CONFIG"]).exists():
        input("Setup did not finish. Press Enter to close.")
        sys.exit(1)

import web  # noqa: E402

web._auto_scan()
host = web.CFG.get("host", "127.0.0.1")
port = int(web.CFG.get("port", 8080))
threading.Timer(1.5, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
print(f"\nBalance check is running at http://127.0.0.1:{port}. Keep this window open while you work; close it to stop.\n")
try:
    from waitress import serve
    serve(web.app, host=host, port=port, threads=8)
except ImportError:
    web.app.run(host=host, port=port)
