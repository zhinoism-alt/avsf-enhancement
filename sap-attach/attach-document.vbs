' Attaches the same file to a list of vendors in SAP (XK02 > Services for Object > Create attachment).
' Built from a recording. SAP GUI must be open and logged in. It uses your first SAP window,
' so leave SAP alone while it runs. Results are written to attach-log.txt next to this script.
Option Explicit

' ===== Settings =====
Dim VENDOR_FILE, ATTACH_FOLDER, ATTACH_NAME, DRY_RUN, START_ROW, MAX_VENDORS, MAX_FAILS_IN_ROW, WAIT_SECS
VENDOR_FILE = "C:\Users\290158\Downloads\Scripts\Attachment Script.xlsx"   ' Excel file, vendor numbers in column A of the first sheet
ATTACH_FOLDER = "C:\Users\290158\Downloads\Scripts"
ATTACH_NAME = "RE_ Suspensiones hasta nuevo aviso.msg"
DRY_RUN = False           ' True = open each vendor and its attachment list but attach NOTHING
START_ROW = 4             ' first Excel row to process (rows 2 and 3 were already done in the tests)
MAX_VENDORS = 0           ' 0 = all vendors in the Excel file; use a small number to test
MAX_FAILS_IN_ROW = 3      ' stop after this many failures in a row
WAIT_SECS = 120           ' how long to wait for slow SAP screens
' ====================

' tidy the settings: remove stray quote marks (e.g. from "Copy as path") and trailing backslashes
VENDOR_FILE = CleanPath(VENDOR_FILE)
ATTACH_FOLDER = CleanPath(ATTACH_FOLDER)
ATTACH_NAME = CleanPath(ATTACH_NAME)

Function CleanPath(p)
  p = Trim(Replace(p, Chr(34), ""))
  Do While Right(p, 1) = "\"
    p = Left(p, Len(p) - 1)
  Loop
  CleanPath = p
End Function

Const LIST_SHELL = "wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell"

Dim fso, session, application, connection, SapGuiAuto, logPath, gErr, runLabel
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "attach-log.txt")
gErr = ""

Sub Log(msg)
  Dim f
  Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  " & msg
  f.Close
End Sub

Sub Trace(m)
  Log "      " & m
End Sub

Function Find(id)
  Dim o
  Set o = Nothing
  On Error Resume Next
  Set o = session.findById(id, False)
  If Err.Number <> 0 Then
    Err.Clear
    Set o = Nothing
  End If
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

Function Press(id)
  Press = False
  On Error Resume Next
  session.findById(id).press
  If Err.Number <> 0 Then
    gErr = "press " & id & ": " & Err.Description
    Err.Clear
  Else
    Press = True
  End If
  On Error GoTo 0
End Function

Function SetText(id, value)
  SetText = False
  On Error Resume Next
  session.findById(id).text = value
  If Err.Number <> 0 Then
    gErr = "set text " & id & ": " & Err.Description
    Err.Clear
  Else
    SetText = True
  End If
  On Error GoTo 0
End Function

Function SetChecked(id, value)
  SetChecked = False
  On Error Resume Next
  session.findById(id).selected = value
  If Err.Number <> 0 Then
    gErr = "tick " & id & ": " & Err.Description
    Err.Clear
  Else
    SetChecked = True
  End If
  On Error GoTo 0
End Function

Function SendKey(id, key)
  SendKey = False
  On Error Resume Next
  session.findById(id).sendVKey key
  If Err.Number <> 0 Then
    gErr = "key " & key & " on " & id & ": " & Err.Description
    Err.Clear
  Else
    SendKey = True
  End If
  On Error GoTo 0
End Function

Function StatusError()
  Dim sb
  StatusError = ""
  On Error Resume Next
  Set sb = session.findById("wnd[0]/sbar")
  If Err.Number = 0 Then
    If sb.messageType = "E" Then StatusError = sb.text
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function PopupTitle(n)
  Dim w
  PopupTitle = ""
  Set w = Find("wnd[" & n & "]")
  If Not w Is Nothing Then
    On Error Resume Next
    PopupTitle = w.text
    Err.Clear
    On Error GoTo 0
  End If
End Function

' lists the buttons of the title-bar toolbar (diagnostics, written to the log once per vendor)
Sub DumpToolbar()
  Dim sh, n, i
  On Error Resume Next
  Set sh = session.findById("wnd[0]/titl/shellcont/shell")
  n = sh.GetButtonCount
  If Err.Number <> 0 Then
    Trace "toolbar dump failed: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Sub
  End If
  Trace "title toolbar has " & n & " buttons"
  For i = 0 To n - 1
    Trace "   button " & i & ": id=" & sh.GetButtonId(i) & "  text=" & sh.GetButtonText(i) & "  type=" & sh.GetButtonType(i)
  Next
  Err.Clear
  On Error GoTo 0
End Sub

Function RowCountOf(grid)
  Dim n
  On Error Resume Next
  n = grid.RowCount
  If Err.Number <> 0 Then
    n = -1
    Err.Clear
  End If
  On Error GoTo 0
  RowCountOf = n
End Function

Function TitleText()
  TitleText = ""
  On Error Resume Next
  TitleText = session.findById("wnd[0]").text
  If Err.Number <> 0 Then
    Err.Clear
    TitleText = session.findById("wnd[0]/titl").text
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function ScreenNo()
  ScreenNo = ""
  On Error Resume Next
  ScreenNo = session.info.screenNumber
  Err.Clear
  On Error GoTo 0
End Function

' Closes one popup SAFELY. It never confirms a question with "Yes" / Enter, so it cannot log you off or save.
Function ClosePopup(n)
  Dim title
  ClosePopup = False
  If Find("wnd[" & n & "]") Is Nothing Then Exit Function
  title = PopupTitle(n)
  If InStr(LCase(title), "attachment list") > 0 And Not Find("wnd[" & n & "]/tbar[0]/btn[0]") Is Nothing Then
    Press "wnd[" & n & "]/tbar[0]/btn[0]"
  ElseIf Not Find("wnd[" & n & "]/usr/btnSPOP-OPTION2") Is Nothing Then
    Press "wnd[" & n & "]/usr/btnSPOP-OPTION2"
  Else
    On Error Resume Next
    session.findById("wnd[" & n & "]").Close
    Err.Clear
    On Error GoTo 0
  End If
  Trace "closed popup: " & title
  WScript.Sleep 700
  ClosePopup = True
End Function

' close popups left over from a previous step (innermost first)
Sub ClosePopups()
  Dim i
  For i = 4 To 1 Step -1
    ClosePopup i
  Next
End Sub

Function IsVendorNumber(s)
  Dim re
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = "^[0-9]{1,10}$"
  IsVendorNumber = re.Test(s)
End Function

Function ReadVendors()
  Dim xl, wb, ws, lastRow, r, v, d
  Set d = CreateObject("Scripting.Dictionary")
  Set xl = CreateObject("Excel.Application")
  xl.Visible = False
  xl.DisplayAlerts = False
  Set wb = xl.Workbooks.Open(VENDOR_FILE, 0, True)
  Set ws = wb.Worksheets(1)
  lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
  For r = 1 To lastRow
    v = Trim(CStr(ws.Cells(r, 1).Value))
    If IsVendorNumber(v) Then d.Add r, Right("0000000000" & v, 10)
  Next
  wb.Close False
  xl.Quit
  Set ReadVendors = d
End Function

' Returns True when the file was attached (or, in dry run, when the attachment list opened).
Function DoVendor(lifnr, ByRef msg)
  Dim grid, before, after, t, ready, i, title, ok3, errText, initScreen, dumped
  DoVendor = False
  msg = ""
  If Find("wnd[0]") Is Nothing Then msg = "SAP session is no longer available (logged off?)": Exit Function
  ClosePopups

  Trace "opening XK02 for vendor " & lifnr
  ' 1. XK02 start screen, vendor number, Address view (as recorded)
  If Not SetText("wnd[0]/tbar[0]/okcd", "/nXK02") Then msg = gErr: Exit Function
  If Not SendKey("wnd[0]", 0) Then msg = gErr: Exit Function
  If WaitFor("wnd[0]/usr/ctxtRF02K-LIFNR", 60) Is Nothing Then msg = "XK02 start screen did not appear": Exit Function
  initScreen = ScreenNo()
  If Not SetChecked("wnd[0]/usr/chkRF02K-D0110", True) Then msg = gErr: Exit Function
  If Not SetText("wnd[0]/usr/ctxtRF02K-LIFNR", lifnr) Then msg = gErr: Exit Function
  If Not SendKey("wnd[0]", 0) Then msg = gErr: Exit Function

  ' 2. wait for the vendor screen (can take ~30 s). SAP may show the "Attachment list" popup by itself.
  ready = False
  t = Timer
  Do While Timer - t < WAIT_SECS
    If Not Find("wnd[1]") Is Nothing Then
      title = PopupTitle(1)
      ClosePopup 1
      If InStr(LCase(title), "attachment list") = 0 Then msg = "Unexpected popup while opening the vendor: " & title: Exit Function
    End If
    If StatusError() <> "" Then msg = "SAP message: " & StatusError(): Exit Function
    If InStr(LCase(TitleText()), "address") > 0 And Not Find("wnd[0]/titl/shellcont/shell") Is Nothing And Find("wnd[1]") Is Nothing Then
      ready = True
      Exit Do
    End If
    WScript.Sleep 500
  Loop
  If Not ready Then msg = "Vendor Address screen did not open within " & WAIT_SECS & " s (window title: " & TitleText() & ")": Exit Function

  ' 3. Services for Object > Attachment list (the toolbar can load slowly, so retry on errors)
  Trace "vendor screen is open (" & TitleText() & ", " & Int(Timer - t) & " s), opening Services for Object"
  ok3 = False
  errText = ""
  dumped = False
  t = Timer
  Do While Timer - t < WAIT_SECS
    ' SAP may already have opened the attachment list by itself
    If Not Find(LIST_SHELL) Is Nothing Then
      Trace "attachment list is already open"
      ok3 = True
      Exit Do
    End If
    errText = ""
    On Error Resume Next
    session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
    If Err.Number <> 0 Then
      errText = "toolbox button: " & Err.Description
      Err.Clear
    Else
      session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
      If Err.Number <> 0 Then
        errText = "attachment menu item: " & Err.Description
        Err.Clear
      End If
    End If
    On Error GoTo 0
    If errText = "" Then
      ok3 = True
      Exit Do
    End If
    Trace "Services for Object not ready yet: " & errText
    If Not dumped Then
      DumpToolbar
      dumped = True
    End If
    WScript.Sleep 3000
  Loop
  If Not ok3 Then msg = "Services for Object menu: " & errText: Exit Function
  Trace "waiting for the attachment list"
  Set grid = WaitFor(LIST_SHELL, WAIT_SECS)
  If grid Is Nothing Then msg = "Attachment list did not open within " & WAIT_SECS & " s": Exit Function
  before = RowCountOf(grid)

  If DRY_RUN Then
    msg = "dry run: attachment list opened (" & before & " existing attachments)"
    DoVendor = True
  Else
    ' 4. Create > Create attachment (from PC) > SAP's own "Import file" dialog
    Trace "creating the attachment"
    On Error Resume Next
    grid.pressToolbarContextButton "%ATTA_CREATE"
    grid.selectContextMenuItem "%GOS_PCATTA_CREA"
    If Err.Number <> 0 Then
      msg = "Create attachment menu: " & Err.Description
      Err.Clear
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    If WaitFor("wnd[1]/usr/ctxtDY_PATH", WAIT_SECS) Is Nothing Then msg = "The Import file dialog did not appear": Exit Function
    If Not SetText("wnd[1]/usr/ctxtDY_PATH", ATTACH_FOLDER) Then msg = gErr: Exit Function
    If Not SetText("wnd[1]/usr/ctxtDY_FILENAME", ATTACH_NAME) Then msg = gErr: Exit Function
    If Not Press("wnd[1]/tbar[0]/btn[0]") Then msg = gErr: Exit Function

    ' 5. wait until the attachment list is back and the file dialog is gone
    ready = False
    t = Timer
    Do While Timer - t < WAIT_SECS
      If Find("wnd[1]/usr/ctxtDY_PATH") Is Nothing And Not Find(LIST_SHELL) Is Nothing Then
        ready = True
        Exit Do
      End If
      If Not Find("wnd[2]") Is Nothing Then msg = "Import file error popup: " & PopupTitle(2): Exit Function
      WScript.Sleep 500
    Loop
    If Not ready Then msg = "Attachment was not confirmed within " & WAIT_SECS & " s": Exit Function
    after = RowCountOf(Find(LIST_SHELL))
    If before >= 0 And after >= 0 And after <= before Then msg = "list still shows " & after & " attachments (was " & before & ")": Exit Function
    msg = "attached (" & before & " -> " & after & " attachments)"
    DoVendor = True
  End If

  ' 6. close the list and leave XK02 (attachments are stored as soon as they are created)
  Trace "closing the list and leaving XK02"
  ClosePopups
  If SetText("wnd[0]/tbar[0]/okcd", "/n") Then SendKey "wnd[0]", 0
  WScript.Sleep 1500
  ClosePopups
End Function

' ---- attach to the running SAP session ----
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
  Log "ERROR: SAP GUI is not running."
  MsgBox "SAP GUI is not running. Open SAP and log in first.", 16, "Attach document"
  WScript.Quit 1
End If
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children(0)
Set session = connection.Children(0)
If Err.Number <> 0 Then
  Log "ERROR: could not attach to a logged-in SAP session."
  MsgBox "Could not attach to a logged-in SAP session.", 16, "Attach document"
  WScript.Quit 1
End If
On Error GoTo 0
session.findById("wnd[0]").maximize

Dim vendors, k, done, okCount, failCount, fails, msg, started, stoppedEarly
done = 0: okCount = 0: failCount = 0: fails = 0: stoppedEarly = False
If Not fso.FileExists(VENDOR_FILE) Then
  MsgBox "Vendor file not found: " & VENDOR_FILE, 16, "Attach document"
  WScript.Quit 1
End If
Set vendors = ReadVendors()
If DRY_RUN Then runLabel = "DRY RUN" Else runLabel = "REAL RUN"
Log "===== " & runLabel & ": " & vendors.Count & " vendors in " & VENDOR_FILE & " | file " & ATTACH_FOLDER & "\" & ATTACH_NAME & " ====="

For Each k In vendors.Keys
  If k >= START_ROW Then
    If MAX_VENDORS > 0 And done >= MAX_VENDORS Then Exit For
    done = done + 1
    started = Timer
    msg = ""
    If DoVendor(vendors(k), msg) Then
      okCount = okCount + 1
      fails = 0
      Log "row " & k & "  vendor " & vendors(k) & "  OK   " & msg & "  (" & Int(Timer - started) & " s)"
    Else
      failCount = failCount + 1
      fails = fails + 1
      Log "row " & k & "  vendor " & vendors(k) & "  FAILED   " & msg
      If InStr(msg, "no longer available") > 0 Then
        stoppedEarly = True
        Log "Stopped: the SAP session is gone."
        Exit For
      End If
      ClosePopups
      If fails >= MAX_FAILS_IN_ROW Then
        stoppedEarly = True
        Log "Stopped after " & fails & " failures in a row."
        Exit For
      End If
    End If
  End If
Next

Log "===== finished: " & okCount & " ok, " & failCount & " failed" & IIf2(stoppedEarly) & " ====="
MsgBox runLabel & " finished." & vbCrLf & okCount & " ok, " & failCount & " failed" & IIf2(stoppedEarly) & "." & vbCrLf & "Details: " & logPath, 64, "Attach document"

Function IIf2(flag)
  If flag Then IIf2 = " (stopped early)" Else IIf2 = ""
End Function
