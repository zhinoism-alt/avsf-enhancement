' Exports the vendor request report (ZMMVEND_DIS) to Excel for the AVSF dashboard.
' Built from a recording of the manual export.
' SAP GUI must be open and logged in, and the PC unlocked.
Option Explicit

' ===== Settings =====
Dim OUT_DIR, OUT_FILE, MONTHS_BACK, TCODE, DEBUG_POPUPS
OUT_DIR = "C:\Users\290158\Documents\avsf-inbox"   ' the folder connected in the dashboard
OUT_FILE = "export.xlsx"
MONTHS_BACK = 3          ' export from the 1st of this many months ago until today
TCODE = "ZMMVEND_DIS"
DEBUG_POPUPS = True      ' True = a message box after each step (use for the first runs); then set False
' ====================

Dim fso, sh, session, application, connection, SapGuiAuto, logPath, outPath, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
If Not fso.FolderExists(OUT_DIR) Then fso.CreateFolder OUT_DIR
logPath = fso.BuildPath(OUT_DIR, "export-log.txt")
outPath = fso.BuildPath(OUT_DIR, OUT_FILE)

Sub Log(msg)
  Dim f
  Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  " & msg
  f.Close
End Sub

Sub StepDone(msg)
  Log msg
  If DEBUG_POPUPS Then MsgBox msg, 64, "AVSF export"
End Sub

Sub Fail(msg)
  Log "ERROR: " & msg
  If DEBUG_POPUPS Then MsgBox "ERROR: " & msg, 16, "AVSF export"
  WScript.Quit 1
End Sub

Function Ymd(d)
  Ymd = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2)
End Function

Function Find(id)
  Dim o
  Set o = Nothing
  On Error Resume Next
  Set o = session.findById(id, False)
  If Err.Number <> 0 Then Err.Clear: Set o = Nothing
  On Error GoTo 0
  Set Find = o
End Function

Function WaitFor(id, seconds)
  Dim i, o
  Set o = Nothing
  For i = 1 To seconds * 2
    Set o = Find(id)
    If Not o Is Nothing Then Exit For
    WScript.Sleep 500
  Next
  Set WaitFor = o
End Function

Sub PickDate(fieldId, ymdValue, firstVisible)
  Dim cal
  session.findById(fieldId).setFocus
  session.findById(fieldId).caretPosition = 0
  session.findById("wnd[0]").sendVKey 4
  Set cal = WaitFor("wnd[1]/usr/cntlCONTAINER/shellcont/shell", 10)
  If cal Is Nothing Then Fail "The calendar popup did not open for " & fieldId
  cal.focusDate = ymdValue
  If firstVisible <> "" Then cal.firstVisibleDate = firstVisible
  cal.selectionInterval = ymdValue & "," & ymdValue
End Sub

' ---- attach to the running SAP session ----
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then Err.Clear: Fail "SAP GUI is not running. Open SAP and log in first."
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children(0)
Set session = connection.Children(0)
If Err.Number <> 0 Then Err.Clear: Fail "Could not attach to a logged-in SAP session."
On Error GoTo 0

Dim fromDate, toDate
fromDate = DateSerial(Year(Date), Month(Date) - MONTHS_BACK, 1)
toDate = Date
StepDone "Attached to SAP. Date range " & Ymd(fromDate) & " to " & Ymd(toDate)

session.findById("wnd[0]").maximize
' always start from a clean main screen, then open the transaction directly
session.findById("wnd[0]/tbar[0]/okcd").text = "/n" & TCODE
session.findById("wnd[0]").sendVKey 0
If WaitFor("wnd[0]/usr/ctxtZTRC_VENDORS-REQ_DATE", 20) Is Nothing Then Fail "Selection screen of " & TCODE & " did not appear."

PickDate "wnd[0]/usr/ctxtZTRC_VENDORS-REQ_DATE", Ymd(fromDate), Ymd(fromDate)
PickDate "wnd[0]/usr/ctxtZTRC_VEND_CCODE-CERDT", Ymd(toDate), ""
session.findById("wnd[0]").sendVKey 0
StepDone "Dates entered."

' the recording went straight to the result list; if it is not there yet, press Execute (F8)
If WaitFor("wnd[0]/tbar[1]/btn[43]", 6) Is Nothing Then session.findById("wnd[0]").sendVKey 8
If WaitFor("wnd[0]/tbar[1]/btn[43]", 120) Is Nothing Then Fail "The result list (Excel button) did not appear."
StepDone "Result list is open."

' ---- export to Excel ----
If fso.FileExists(outPath) Then fso.DeleteFile outPath, True
session.findById("wnd[0]/tbar[1]/btn[43]").press

' format popup (SAP skips it when "Always use selected format" is ticked)
If Not WaitFor("wnd[1]/tbar[0]/btn[0]", 5) Is Nothing Then
  If Find("wnd[1]/usr/ctxtDY_PATH") Is Nothing Then session.findById("wnd[1]/tbar[0]/btn[0]").press
End If

' SAP's own save dialog: Directory + File Name, then Generate
If WaitFor("wnd[1]/usr/ctxtDY_PATH", 20) Is Nothing Then Fail "The SAP save dialog did not appear (expected Directory / File Name fields)."
session.findById("wnd[1]/usr/ctxtDY_PATH").text = OUT_DIR & "\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = OUT_FILE
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 1000
' if the file could not be deleted beforehand, the dialog stays open with Replace available
If Not Find("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
  If Not Find("wnd[1]/tbar[0]/btn[11]") Is Nothing Then session.findById("wnd[1]/tbar[0]/btn[11]").press
End If

Dim i
For i = 1 To 120
  If fso.FileExists(outPath) Then
    If fso.GetFile(outPath).Size > 0 Then Exit For
  End If
  WScript.Sleep 500
Next
If Not fso.FileExists(outPath) Then Fail "The Excel file was not saved to " & outPath

' possible confirmation popup after saving (recorded as a second OK press)
If Not WaitFor("wnd[1]/tbar[0]/btn[0]", 5) Is Nothing Then session.findById("wnd[1]/tbar[0]/btn[0]").press
StepDone "Saved " & outPath & " (" & fso.GetFile(outPath).Size & " bytes)"

' SAP opens the file in Excel after saving: close just that workbook so it is not locked
Dim xl, wb
On Error Resume Next
Set xl = GetObject(, "Excel.Application")
If Err.Number = 0 Then
  For Each wb In xl.Workbooks
    If LCase(wb.Name) = LCase(OUT_FILE) Then wb.Close False
  Next
End If
Err.Clear
On Error GoTo 0

session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0
Log "Done."
If DEBUG_POPUPS Then MsgBox "Export finished.", 64, "AVSF export"
