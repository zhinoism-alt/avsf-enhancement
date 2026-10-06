# Attach one document to many vendors (XK02)

`attach-document.vbs` repeats what you did by hand for one vendor: XK02 → vendor → Services for Object →
Attachment list → Create attachment → file → leave without saving. It reads vendor numbers from Excel and
writes one line per vendor to `attach-log.txt` (next to the script).

## Settings (top of the script)
| Setting | Meaning |
|---|---|
| `VENDOR_FILE` | Excel file; vendor numbers in column A of the first sheet (header row is ignored) |
| `ATTACH_FOLDER` / `ATTACH_NAME` | the file to attach |
| `DRY_RUN` | `True` opens each vendor and its attachment list but attaches nothing |
| `MAX_VENDORS` | how many vendors to process (0 = all) |
| `START_ROW` | continue from this Excel row after an interruption (see the log for the last row done) |
| `MAX_FAILS_IN_ROW` | stops after this many failures in a row |
| `WAIT_SECS` | patience for slow SAP screens (it can take ~30 s per vendor) |

## How to roll it out
1. Put your vendor numbers in the Excel file and set `VENDOR_FILE`. Close the Excel file.
2. **Dry run** with `DRY_RUN = True` and `MAX_VENDORS = 3`: log in to SAP and double-click the script.
   The log must show OK for each vendor.
3. **Real test:** `DRY_RUN = False`, `MAX_VENDORS = 3`. Check the attachments in SAP.
4. **Everything:** `MAX_VENDORS = 0`. About 50 vendors at ~1 minute each; leave SAP alone while it runs.

If it stops, open `attach-log.txt`, fix the cause, set `START_ROW` to the row after the last OK vendor and run
again. Vendors are not skipped automatically, so re-running the same rows would attach the file twice.

Check with your controls team that mass-attaching to vendor records is allowed; every attachment is audited in SAP.

# Payment block + email (`payment-block.vbs`)

Excel (first sheet, row 1 = header): **A** vendor number, **B** PBA (the block to set, e.g. `A`; **blank = remove block A**),
**C** company code (blank = `DEFAULT_CC`, 1505).

Per vendor it opens XK02 (Address + company code Payment transactions), attaches the email on the Address screen,
presses Enter to reach the Payment transactions screen, reads the Payment block and applies these rules:
- wanted `A`, vendor has no block → sets `A` and saves; already `A` → nothing.
- wanted blank, vendor has `A` → removes it and saves; no block → nothing.
- any other block in place → left unchanged and listed in the final message ("left unchanged on purpose").
- any error (SAP message, popup, vendor not in the company code, Save not confirmed) → `FAILED` in `payment-block-log.txt`
  and listed in the message box at the end.

First run: `DRY_RUN = True`, `MAX_VENDORS = 1` (reads and reports only). Then `DRY_RUN = False` for the one test vendor,
then `MAX_VENDORS = 0` for the whole list.
