# SAP scripts: how to run them (one page)

These scripts do repetitive vendor work in SAP for you: **attach a file to many vendors** (`attach-document.vbs`) and
**set or remove a payment block + attach a file** (`payment-block.vbs`). They use **your own SAP login**: whatever they
change is recorded under your user, exactly as if you had done it by hand.

## One-time setup (about 5 minutes)
1. **Copy this whole folder** to your PC, for example `C:\SAP-scripts` (not inside a zip, not on a read-only drive).
   If the files came from email/Teams: right-click each `.vbs` file, **Properties**, tick **Unblock**, OK.
2. **SAP scripting:** in SAP press **Alt+F12**, choose **Options > Accessibility & Scripting > Scripting**. Tick
   **Enable scripting** and untick both **Notify when a script...** boxes. (If it is greyed out, ask IT.)
3. **Log on to SAP in English** (language `EN` on the SAP logon screen). The scripts recognise English screens.
4. **Excel** must be installed. Close the vendor Excel file before running the script.

## Prepare the vendor list (Excel, first sheet, row 1 = header)
- `attach-document.vbs`: column A = vendor number. Optional column with header **File** = full path of the file for that vendor.
- `payment-block.vbs`: A = vendor number, B = **PBA** (`A` to block, blank to remove block A), C = company code, optional **File**.
- See the `examples` folder. Save the list in this folder with the name shown at the top of the script (`VENDOR_FILE`),
  or change that line.

## First run: check everything without touching SAP
1. Open the script in Notepad and change `CHECK_ONLY = False` to `CHECK_ONLY = True`, save.
2. Log on to SAP, leave it on the main screen, double-click the script.
3. A message tells you what to fix, or "All checks passed" with your SAP user, the number of vendors and any missing files.

## Run it for real
1. Set `CHECK_ONLY = False`. For a first try also set `DRY_RUN = True` and `MAX_VENDORS = 1` (looks only, changes nothing).
2. Double-click the script and **leave SAP alone** (no clicking, keep SAP in front) until the final message box appears.
3. Read the message box and `...-log.txt` (next to the script). Every problem is listed there.
4. When the dry run looks right: `DRY_RUN = False`, `MAX_VENDORS = 0` (all vendors).

## Good to know
- Only one copy may run at a time. If it says another copy is running: Ctrl+Shift+Esc, end every `wscript.exe`, start again.
- The same vendor is skipped if it already got the same thing **today** (list in `...-done.txt`). Delete that file to repeat.
- Changing vendor master data should be approved by your controls team.
