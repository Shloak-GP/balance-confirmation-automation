# Trial on the company server (Windows, e.g. AI-TEST)

The tool runs on the server. People open it in a browser at http://SERVERNAME:8080. No Outlook on the server:
replies reach it as files in a folder.

1. Install Python 3.11+ (tick "Add python.exe to PATH"). Unzip this app to `D:\BalanceCheck`.
2. `cd D:\BalanceCheck` then `pip install -r requirements.txt`
3. Create these folders (any drive the server user can write to):
   - `D:\BalanceCheck_data\Lists`    his consolidated list lives HERE (copy the Excel file in)
   - `D:\BalanceCheck_data\Replies`  reply mails and PDFs go here (a shared folder is fine)
   - `D:\BalanceCheck_data\Reports`  MIS files made by the tool
   - `D:\BalanceCheck_data\Data`     the tool's own database and backups
4. Copy `config.example.json` to `config.json`, then set:
   - `consolidated_list`: `D:/BalanceCheck_data/Lists/Consolidated List.xlsx`
   - `sources`: `[{"name": "Shared folder", "path": "D:/BalanceCheck_data/Replies"}]`
   - `output_folder`, `data_folder`, `write_mode`: `in_place`
   - `host`: `0.0.0.0`, `port`: `8080`, `auto_scan_minutes`: `10`
   - `report_link_prefix`: `""` (or the shared path the boss can open), `boss_email`
   - Leave `outlook_folder` empty (Outlook buttons do not work on a server).
   Use `/` or `\\` in paths, never single `\`.
5. Run `run_server.bat`. Open http://localhost:8080/health on the server: every line must say ok.
6. Windows Firewall: allow inbound TCP 8080 for the office network only.
7. Keep it running: Task Scheduler, "At startup", action `run_server.bat`, "Run whether user is logged on or not".

Working day: replies land in `Replies` (the regions save them; for mail sent straight to him he runs
`pull_outlook.py` on his own PC with `outlook_save_to` pointing at the shared Replies folder). He opens
the page, checks replies, and downloads the Excel/MIS from the page (or opens them from the shared folder).

Known gaps in this trial: no login (anyone who can reach the address can use it), one server process only,
untested on Windows Server.
