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

## 2. Record the SAP export once (needed to build the script)
SAP scripting is already enabled for your user. In SAP GUI Options → Accessibility & Scripting →
Scripting, **untick "Notify when a script attaches to SAP GUI"** and **"…opens a connection"**,
otherwise a confirmation window blocks the script every time.

Then record the steps:
1. In SAP press **Alt+F12 → Script Recording and Playback**, choose a save path/file name
   (for example `export-vendors.vbs`) and click the **Record** (red dot) button.
2. Do the export exactly as you do today: run `ZMMVEND_DIS`, set the date range, execute,
   click the Excel icon, choose "Select from all available formats" → **10 Excel (XLSX)**,
   confirm, type the file name, save.
3. Click **Stop**. Send me the `.vbs` file (you can paste it here). I will turn it into a script that:
   attaches to the open SAP session, sets the date range automatically
   (first day of 3 months ago → today), and saves to `avsf-inbox\export.xlsx`.

## 3. Schedule it
Windows Task Scheduler → Create Task → trigger weekdays 08:05 → action
`wscript.exe "C:\path\to\export-vendors.vbs"`; tick "Run only when user is logged on".
SAP must be open and logged in at that time (the script can also start SAP if your
logon uses single sign-on; tell me your system ID and client).

## Date range
Use a window of about 3 months, not "first of the month". Requests still open from earlier
months would otherwise never be refreshed and would stay "Processing" in the dashboard.
