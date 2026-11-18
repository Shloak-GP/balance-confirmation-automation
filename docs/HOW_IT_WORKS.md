# balance·check — balance confirmation review tool (first version)

A small web page for the balance-confirmation owner. It reads new replies from the regions' shared
folder, suggests CLEAR / DISPUTE, he confirms each one with one click, and it writes his
consolidated list and MIS in his own Excel template. No AI model, no paid API, no internet needed.

**Status (Nov 17):** tested on FAKE data and on 3 real replies (.msg) plus the real ERP dump. The rules that read
replies must still be checked against all ~17 real replies and the 2 PDFs before anyone relies on it.

## What it does
1. **Collect (several sources)** – `sources` in config lists every place replies come from (the regions'
   shared folder, a folder his mails land in, ...). All are read-only. The same reply found in more
   than one place (an RM saved the PDF and it is also in his mail; or two RMs saved it) is recognised by
   its *content* (the PDF itself, or subject + body) and becomes ONE item that remembers every source.
   The **Mail vs folder** page lists what is in both and what is only in one, so he can chase what
   an RM forgot to save. Original single-folder description: every time Home opens (or "Look for new replies"), it reads new `.pdf`, `.eml`,
   `.msg`, `.txt` files under `shared_folder`. Read-only: nothing in the folder is moved or edited.
   A byte-identical copy (same reply saved by two RMs) is skipped.
2. **Find the customer** – the customer No. from the list is looked for in the file name, then the
   mail/PDF text. It is accepted only if the name also matches; otherwise he picks the customer.
3. **Suggest** – fixed rules: "no changes / we confirm" = CLEAR with our balance; a different amount
   = DISPUTE; any condition ("except", "subject to", "but", "pending") or no readable text
   (scans, handwriting) = he types the amount. Any difference, even paise, is a dispute.
4. **Confirm** – one reply per screen, our ERP balance vs the customer's balance, CLEAR/DISPUTE in
   colour. Nothing is written until he presses "Correct, save". "Saved" lets him reopen and fix.
5. **Write** – default `write_mode: "in_place"`: updates his real consolidated list, but only the
   tool's own columns (BALANCE AS PER Customer, DIFFERENCE, Clear / Dispute, Remarks, RECVD ON,
   Balance Confirmation Status) for replies he saved, plus the MIS counts. **He can edit every other
   column in Excel.** If the file is open in Excel, nothing is written; a banner says so and the
   tool writes as soon as it is closed. A backup is taken before every write (`data_folder/backups`,
   last 30). Changes to the tool's six columns go through "Saved → Fix" so the two never disagree.
   `write_mode: "copy"` instead never touches his file and writes a dated copy to `output_folder`.
6. **MIS mail** – "Open email in Outlook" opens a new Outlook mail (mailto link) with To, subject,
   per-region summary and the **MIS file's** shared-drive location filled in; he presses Send. A mailto
   link cannot attach files, so "Draft with MIS attached" is also offered (downloads an Outlook
   draft .eml with the MIS attached). A draft that appears straight in his Outlook Drafts folder
   needs Microsoft 365 / Graph mail access from IT (not built).

Every save is logged (who, when, what) in `data_folder/balance_confirmations.db`.

## New in v2
- **Checks** page: reads his whole list and flags rows that do not add up (Clear with a difference, Dispute with none,
  difference != ERP minus customer, huge difference, opposite sign, Received without a status, date problems, duplicates)
  plus MIS typed numbers that differ from the list. It never changes a row by itself; each flag has buttons
  (fix it / it is correct, keep). Every action is logged.
- **New period** page: upload the half-yearly ERP dump -> a new consolidated list in his template. Region comes from the
  sales region code, group names are mapped, Priority / Email id / team names are carried over from the previous list
  by customer No. (the dump does not contain them), new customers are reported. His current list is never touched.
- Amount reading fixed for real wording (`Rs. 1,67,75,671/-`, `-3,452,112.00`, phone numbers, quoted older mails, period totals).
- `.msg` is read with `olefile` (pure Python; `extract-msg` would not install).
- His MIS formulas (P2, Pending, %) are never overwritten.
- **Open Excel / Open MIS in Excel** buttons: open the file on his computer in Excel (his real list in its own folder; the MIS in the reports folder). Works only when the tool runs on his own computer (host 127.0.0.1). On a server it shows the file's location with a Copy button instead; set `open_on_server: true` only if the server IS his computer.

## Install on the company server (AI-TEST or similar; CPU only)
Recommended: run on Windows directly (simplest for reaching `\\server\share` paths).
```
py -m pip install -r requirements.txt
copy config.example.json config.json      # then edit the paths
py web.py                                  # opens on port 8080
```
- `shared_folder`: the regions' folder (UNC path like `\\\\fileserver\\Finance\\BalConf\\Replies` is fine).
- `consolidated_list`: his list (the ERP export, his template). The tool never edits this file.
- `output_folder`: where the Excel he sends is written (shared drive or SharePoint-synced folder).
- `data_folder`: the tool's own database and page images; keep it on the server's local disk.
- `division_map`: **Customer Posting Group → MIS division. Must come from him.** Unmapped groups
  are shown as a warning on the MIS page and not counted.
- Run it as a service under a service account that can read `shared_folder` and write
  `output_folder` (NSSM or Task Scheduler "at startup"). Open TCP 8080 inbound for the office network.
- He opens `http://<server-name>:8080` in his browser. Add a desktop shortcut for him.

## Tests
```
py -m pip install pytest
py -m pip install -r requirements-dev.txt
py -m pytest -q tests        # 45 tests: rules + full flow on fake data
```

## Known limits (first version)
- Scanned/handwritten replies: the amount is typed by him (free OCR reads handwriting badly).
  Printed scans also need typing for now; adding local OCR for printed scans is a later step
  once real samples show whether it is worth it.
- One user at a time (SQLite). Fine for one owner.
- Login is a single shared `access_password` (empty = no login); anyone on the network who knows the URL and password can open it. Put it behind Windows
  authentication / IT's reverse proxy before wider use.
- `.msg` needs the optional `extract-msg` package.
- Timing differences, partial confirmations and Hindi wording are handled only as "he types it".

## Getting his mail into the tool
The tool reads *folders*; it does not log into Outlook. Ways to fill the "His mail" folder:
1. **Manual (works today, no IT):** he selects mails in Outlook and drags them into that folder
   (creates .msg files; needs `pip install extract-msg`). Slow if there are many.
2. **Power Automate (needs Microsoft 365 + IT):** a flow "when a mail arrives in his Inbox with an
   attachment -> save mail + attachment to the folder". No code from us; IT must allow it.
3. **Microsoft Graph connector (needs Microsoft 365 + IT admin consent):** the tool reads his mailbox
   directly ("Pull from mailbox"). Built and unit-tested; not yet tried on the real tenant. See GRAPH_SETUP.md.
All three feed the same screen; de-duplication works no matter how a reply arrived.

## What goes to the boss
Only the **MIS** (`MIS_Balance_Confirmation_<date>.xlsx`, the MIS sheet alone). The consolidated list with
customer names and balances is never attached or linked in the boss email; "Download full list" is for him only.

## Our balance
Always taken from the consolidated list (BALANCE AS PER ERP), never from the mail. A reply that does not
mention any balance ("we confirm, no changes") is compared against the list's balance.
