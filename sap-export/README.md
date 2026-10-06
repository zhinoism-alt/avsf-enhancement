# Automatic SAP export (ZMMVEND_DIS → Excel)

Goal: every weekday at about 08:05 a script exports the vendor request report to the
inbox folder, and the dashboard (open in Edge) imports it by itself.

## 1. One-time dashboard setup
1. Create a folder, for example `C:\Users\<you>\Documents\avsf-inbox`.
2. Open the dashboard in **Edge** → AVSF Tracker → **📁 Connect SAP folder** → pick that folder.
   When Edge asks, choose **Allow on every visit** so it does not ask again.
3. Let the dashboard start with Windows: Edge menu → Apps → Install this site as an app,
   then Windows Settings → Apps → Startup → turn on the dashboard app.

New `.xlsx` files in that folder that contain a "Vendor Request No." column are imported
within about 2 minutes. Other Excel files are ignored. Each import is logged in the Audit Log
and updates the "SAP data: …" chip.

## 2. Install the export script
1. Copy `export-vendor-requests.vbs` to a folder on your PC, for example `Documents\avsf-inbox\scripts`.
2. Open it in Notepad and check the settings at the top: `OUT_DIR` (the inbox folder you connected in
   the dashboard), `MONTHS_BACK` (default 3) and `TCODE`.
3. In SAP GUI Options → Accessibility & Scripting → Scripting, **untick** "Notify when a script attaches
   to SAP GUI" and "...opens a connection", otherwise a confirmation window blocks the script.
4. **First test:** log in to SAP (any screen) and double-click the script. `DEBUG_POPUPS = True` shows a
   message after each step. Check that `export-<date>.xlsx` appears in the inbox folder and that the dashboard
   imports it within about 2 minutes ("SAP data" chip updates).
5. When it works, change `DEBUG_POPUPS` to `False`.

The script opens `ZMMVEND_DIS` itself (it first goes to the main screen), picks both dates in the SAP calendar
popups like the recording did, clicks the Excel button, fills SAP's own save dialog
(Directory / File Name / Generate), waits for the file, closes the copy SAP opens in Excel and returns to the
main screen. Progress and errors are written to `export-log.txt` in the inbox folder.

## 3. Schedule it
Double-click `schedule-export-task.cmd` (keep it next to the `.vbs`). It creates the Windows task
"AVSF SAP export": weekdays at 08:05, while you are logged in. SAP must be open and logged in and the PC
unlocked at that time. Test it once without waiting: Task Scheduler → "AVSF SAP export" → Run.

Each run saves a new file named `export-YYYYMMDD-HHMMSS.xlsx` (so there is never an existing file to
replace); exports older than 7 days are deleted automatically. Set `DEBUG_POPUPS = True` in the script only
when troubleshooting.

## Date range
Use a window of about 3 months, not "first of the month". Requests still open from earlier
months would otherwise never be refreshed and would stay "Processing" in the dashboard.
