# Balance-confirmation automation

A small web tool for the finance owner of customer balance confirmations. It reads the replies customers
send (mails, `.msg`, PDFs), finds the customer in the consolidated list, suggests **CLEAR** or **DISPUTE**
by fixed rules, lets him confirm each reply with one click, then writes his consolidated list and the MIS
in his own Excel template. No AI model and no paid API: plain rules on a CPU server.

Rule: CLEAR only if the customer's balance equals our ERP balance. Any difference, even paise, is a
DISPUTE. "No changes / we confirm" is CLEAR. Our balance always comes from the consolidated list.
Nothing is written until he presses **Correct, save**.

Status: tested on fake data and on real reply wording. Graph mailbox and OneDrive reading are built and
unit-tested but not yet tried against the company tenant.

## Layout
| Path | What |
|---|---|
| `web.py` | Flask app (all pages, `/health`) |
| `bc/` | `rules.py` reading amounts + matching customers, `ingest.py` sources + dedup, `workbook.py` Excel/MIS, `checks.py`, `dump.py` ERP dump to new list, `graph.py` / `imap.py` / `outlook.py` mail sources, `store.py` SQLite log |
| `templates/` | HTML pages |
| `sample/` | Empty Excel template + fake data generator |
| `tests/` | 45 pytest tests, fake data only |
| `docs/` | Workflow doc, diagram (`.drawio`, `.png`), `HOW_IT_WORKS.md` (full feature notes) |
| `SERVER_SETUP.md`, `WINDOWS_SETUP.md`, `GRAPH_SETUP.md`, `IT_REQUEST.md`, `NGROK_TRIAL.md`, `OWNER_GUIDE.md` | Step-by-step guides |

## Quick start on fake data (2 minutes)
Python 3.11+.
```
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
python demo.py                     # builds ./demo_data with fake replies and opens http://127.0.0.1:8080
```

## Tests
```
pip install -r requirements-dev.txt
python -m pytest -q tests          # 45 passed; needs no real data
```

## Configuration
Copy `config.example.json` to `config.json` (git-ignored) and set the paths. The app reads
`config.json` next to `web.py`, or the file named in the `BC_CONFIG` environment variable.
Key settings:
- `sources`: every folder replies come from (shared folder, a folder his mails land in). Read-only. The same
  reply in several places becomes one item.
- `consolidated_list`: his Excel list. `write_mode: "in_place"` updates only the tool's 6 columns + MIS counts
  (backups in `data_folder/backups`); `"copy"` never touches his file and writes a dated copy to `output_folder`.
- `output_folder`, `data_folder`: MIS reports, and the tool's own database.
- `division_map`: Customer Posting Group to MIS division. Must be confirmed by the owner.
- `access_password`: set it on any shared server (empty = no login).

**Secrets never go in a file in this repo.** Use environment variables: `BC_GRAPH_SECRET` (Microsoft Graph),
`BC_IMAP_PASSWORD` (IMAP). `local_overrides.json` (git-ignored, see the `.example`) is only for a personal trial.

## Running on the company server
Option A, Windows server directly (simplest for `\\server\share` paths). Full steps in `SERVER_SETUP.md`:
```
pip install -r requirements.txt
copy config.example.json config.json   & rem edit paths, host 0.0.0.0, port 8080, access_password
run_server.bat                          & rem waitress on port 8080
```
Check `http://localhost:8080/health` (every line ok), open TCP 8080 to the office network only, and keep it
running with Task Scheduler "At startup".

Option B, container (Linux host):
```
mkdir -p data && cp config.docker.json data/config.json   # set access_password, paths are the mounts below
docker compose up -d
```
`docker-compose.yml` mounts the lists, replies and reports folders, runs read-only, non-root, with no capabilities.
Edit the host paths (`/srv/finance/...`) to the real shared folders.

## Microsoft Graph route (reading his mailbox and the OneDrive folder)
Needs Microsoft 365 and an Entra admin. Send `IT_REQUEST.md` to IT; steps are in `GRAPH_SETUP.md`.
1. IT registers an app ("Balance check") with application permissions `Mail.Read`, `Mail.ReadWrite`,
   `Files.Read.All` (or `Sites.Selected`), grants admin consent, and limits it to the one mailbox with an
   Exchange application access policy.
2. Put `graph_tenant`, `graph_client_id`, `graph_mailbox`, `thread_subject` (and `onedrive_link` if used) in
   `config.json`. Put the client secret in `BC_GRAPH_SECRET` (service or container environment).
3. `python graph_test.py` signs in and lists the 5 newest mails (read only). A 403 means a missing permission or consent.
4. In the app: **Pull from mailbox** on Home; **Reply on the thread** on the MIS page creates a draft (never sends).

Without Graph: he drags mails from Outlook into a reply folder, or runs `pull_outlook.py` on his own PC
(Windows + Outlook), or IT sets up a Power Automate flow that saves mails to the folder.

## What is not in this repo
Real customer replies, the ERP consolidated list, `config.json` / `local_overrides.json`, and any password or
secret. `run_real.py` expects a local `real_data/` folder that is never committed.

## Known limits
- Scanned or handwritten replies: he types the amount (no OCR yet).
- One process, SQLite, one owner at a time.
- Login is one shared password; put it behind IT's proxy / Windows auth before wider use.
