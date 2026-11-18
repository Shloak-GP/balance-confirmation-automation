# Setup on his Windows PC (classic Outlook, no Microsoft 365)

1. Install Python 3.11+ from python.org (tick "Add python.exe to PATH").
2. Unzip this folder to `C:\BalanceCheck`.
3. Open Command Prompt there and run: `pip install -r requirements.txt`
4. Copy `config.example.json` to `config.json` and set:
   - `consolidated_list`: full path of his Excel on his local drive (use `\\` or `/`), `write_mode`: `in_place`
   - `sources`: the shared mail and PDF folders (UNC paths such as `\\\\server\\share\\replies` are fine)
   - `host`: `127.0.0.1`, `port`: `8080`, `boss_email`, `thread_subject` (the daily thread's subject)
   - `outlook_folder`: e.g. `Inbox/Balance Confirmation` (shows the "Pull from Outlook" button)
   - `outlook_save_to`: the folder to copy pulled mails into (default: first source)
5. Double-click `start.bat`. The page opens at http://localhost:8080.

Daily: Pull from Outlook, Look for new replies, check each reply, Open Excel, Email MIS (Outlook reply draft, he clicks Send).
Needs classic Outlook open on the same PC. The "new Outlook" is not supported.
Not tested on Windows by the builder: first run on his PC is the real test.
