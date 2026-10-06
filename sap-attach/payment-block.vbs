' Payment block + email for a list of vendors (XK02), built from recordings.
' Excel (first sheet, row 1 = header):  A = vendor number | B = PBA, the payment block to set (A) -
' leave it BLANK to remove block A | C = company code (blank = DEFAULT_CC).
' Rules: block "A" is set only when the vendor has no block; block "A" is removed only when it is block A;
' any other block, or a block that is already as wanted, is left alone (and reported).
' SAP GUI must be open and logged in. It uses your first SAP window, so leave SAP alone while it runs.
' Results: payment-block-log.txt next to this script; a message box at the end lists every problem.
Option Explicit

' ===== Settings =====
Dim VENDOR_FILE, ATTACH_FOLDER, ATTACH_NAME, ATTACH_EMAIL, DEFAULT_CC, DRY_RUN, START_ROW, MAX_VENDORS, MAX_FAILS_IN_ROW, WAIT_SECS, LOAD_SECS, DIALOG_SECS, MENU_SECS, SKIP_DONE
VENDOR_FILE = "C:\Users\290158\Downloads\Scripts\PBA 1505.xlsx"
ATTACH_FOLDER = "C:\Users\290158\Downloads\Scripts"
ATTACH_NAME = "RE_ Suspensiones hasta nuevo aviso.msg"
ATTACH_EMAIL = True       ' True = also attach the email to every vendor
DEFAULT_CC = "1505"       ' company code used when column C is blank
DRY_RUN = True            ' True = look only: opens each vendor, reads its payment block and reports what WOULD change; changes/attaches NOTHING
START_ROW = 2             ' first Excel row to process (row 1 is the header)
MAX_VENDORS = 1           ' 0 = every vendor in the Excel file; keep 1 for the first test
MAX_FAILS_IN_ROW = 3      ' stop after this many failures in a row
SKIP_DONE = True          ' skip vendors already completed with the same block/company code/file (payment-block-done.txt; delete it to start over)
WAIT_SECS = 120           ' how long to wait for slow SAP screens
LOAD_SECS = 0             ' optional pause after the vendor opens (0 = none)
MENU_SECS = 45            ' how long SAP may take to fill a menu
DIALOG_SECS = 150         ' how long SAP may take to open/confirm the Import file dialog
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
Const ZAHLS = "wnd[0]/usr/ctxtLFB1-ZAHLS"   ' Payment block field (company code Payment transactions screen)
Const TOOLBOX = "wnd[0]/shellcont/shell"   ' toolbar with CREATE_ATTA, VIEW_ATTA, ... shown by System > Services for Object

Dim fso, session, application, connection, SapGuiAuto, logPath, doneFile, gErr, runLabel
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "payment-block-log.txt")
doneFile = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "payment-block-done.txt")
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

Function StatusText()
  Dim sb
  StatusText = ""
  On Error Resume Next
  Set sb = session.findById("wnd[0]/sbar")
  If Err.Number = 0 Then StatusText = sb.text
  Err.Clear
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
  n = sh.ButtonCount
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

' native dropdown menus only open in the active window, so bring the toolbox / SAP window to the front
Sub BringSapToFront()
  Dim shl, ttl
  Set shl = CreateObject("WScript.Shell")
  On Error Resume Next
  ttl = ""
  ttl = session.findById("wnd[0]/shellcont").Title
  Err.Clear
  If ttl <> "" Then shl.AppActivate ttl
  Err.Clear
  ttl = TitleText()
  If ttl <> "" Then shl.AppActivate ttl
  Err.Clear
  session.findById("wnd[0]").setFocus
  Err.Clear
  On Error GoTo 0
  WScript.Sleep 500
End Sub

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

' how many copies of this script are running right now (including this one)
Function InstancesRunning()
  Dim wmi, procs, p, n
  n = 0
  On Error Resume Next
  Set wmi = GetObject("winmgmts:\\.\root\cimv2")
  Set procs = wmi.ExecQuery("Select CommandLine from Win32_Process Where Name = 'wscript.exe' Or Name = 'cscript.exe'")
  If Err.Number = 0 Then
    For Each p In procs
      If Not IsNull(p.CommandLine) Then
        If InStr(LCase(p.CommandLine), LCase(WScript.ScriptName)) > 0 Then n = n + 1
      End If
    Next
  End If
  Err.Clear
  On Error GoTo 0
  InstancesRunning = n
End Function

' waits for SAP's Import file dialog (it can take over a minute to appear); closes the
' attachment list if SAP shows that instead. Never confirms anything.
Function WaitForImportDialog(seconds)
  Dim t0
  WaitForImportDialog = False
  t0 = Timer
  Do While Timer - t0 < seconds
    If Not Find("wnd[1]/usr/ctxtDY_PATH") Is Nothing Then
      WaitForImportDialog = True
      Exit Function
    End If
    If Not Find("wnd[1]") Is Nothing Then
      If InStr(LCase(PopupTitle(1)), "attachment list") > 0 Then ClosePopup 1
    End If
    If Not Find("wnd[2]") Is Nothing Then Exit Function
    WScript.Sleep 1000
  Loop
End Function

Function IsVendorNumber(s)
  Dim re
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = "^[0-9]{1,10}$"
  IsVendorNumber = re.Test(s)
End Function

Function ReadVendors()
  Dim xl, wb, ws, lastRow, r, v, d, pbaVal, ccVal
  Set d = CreateObject("Scripting.Dictionary")
  Set xl = CreateObject("Excel.Application")
  xl.Visible = False
  xl.DisplayAlerts = False
  Set wb = xl.Workbooks.Open(VENDOR_FILE, 0, True)
  Set ws = wb.Worksheets(1)
  lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
  For r = 1 To lastRow
    v = Trim(CStr(ws.Cells(r, 1).Value))
    If IsVendorNumber(v) Then
      pbaVal = UCase(Trim(CStr(ws.Cells(r, 2).Value)))
      ccVal = Trim(CStr(ws.Cells(r, 3).Value))
      If ccVal = "" Then ccVal = DEFAULT_CC
      d.Add r, Right("0000000000" & v, 10) & "|" & pbaVal & "|" & ccVal
    End If
  Next
  wb.Close False
  xl.Quit
  Set ReadVendors = d
End Function

' opens the Attachment list (System > Services for Object if the toolbox is not open yet); returns the grid or Nothing
Function OpenList()
  Dim g
  Set g = Nothing
  If Find(TOOLBOX) Is Nothing Then
    On Error Resume Next
    session.findById("wnd[0]/mbar/menu[5]/menu[6]").select
    Err.Clear
    On Error GoTo 0
    If WaitFor(TOOLBOX, 40) Is Nothing Then
      Set OpenList = Nothing
      Exit Function
    End If
  End If
  On Error Resume Next
  session.findById(TOOLBOX).pressButton "VIEW_ATTA"
  Err.Clear
  On Error GoTo 0
  Set g = WaitFor(LIST_SHELL, 40)
  Set OpenList = g
End Function

Function SelectMenu(id)
  SelectMenu = False
  On Error Resume Next
  session.findById(id).select
  If Err.Number <> 0 Then
    gErr = "menu " & id & ": " & Err.Description
    Err.Clear
  Else
    SelectMenu = True
  End If
  On Error GoTo 0
End Function

Function StatusType()
  Dim sb
  StatusType = ""
  On Error Resume Next
  Set sb = session.findById("wnd[0]/sbar")
  If Err.Number = 0 Then StatusType = sb.messageType
  Err.Clear
  On Error GoTo 0
End Function

Function Describe(blk)
  If blk = "" Then
    Describe = "no block"
  Else
    Describe = "block " & blk
  End If
End Function

' Moves to the company code Payment transactions screen and applies the payment block rules.
' Returns True when everything is as wanted (changed, or nothing needed, or left alone on purpose).
Function HandleBlock(desired, ByRef bmsg)
  Dim cur, newVal, t, st, stText, ttl, way, moved
  HandleBlock = False
  bmsg = ""
  desired = UCase(Trim(desired))

  ' Go to the company code Payment transactions screen. SAP can ignore the first way, so three are tried:
  ' 1 = Enter (as in the recording), 2 = the Next screen button, 3 = Goto > Company code data > Payment transactions
  ClosePopups
  If Find(ZAHLS) Is Nothing Then
    For way = 1 To 3
      moved = False
      If way = 1 Then
        Trace "going to the Payment transactions screen: way 1, Enter"
        moved = SendKey("wnd[0]", 0)
      ElseIf way = 2 Then
        Trace "going to the Payment transactions screen: way 2, Next screen button"
        moved = Press("wnd[0]/tbar[1]/btn[8]")
      Else
        Trace "going to the Payment transactions screen: way 3, Goto menu"
        moved = SelectMenu("wnd[0]/mbar/menu[2]/menu[4]/menu[1]")
      End If
      If Not moved Then Trace "way " & way & " could not be used: " & gErr
      t = Timer
      Do While Timer - t < 20
        If Not Find(ZAHLS) Is Nothing Then Exit Do
        If Not Find("wnd[1]") Is Nothing Then
          ttl = PopupTitle(1)
          ClosePopup 1
          If InStr(LCase(ttl), "attachment list") = 0 Then
            bmsg = "unexpected popup while moving to the Payment transactions screen: " & ttl
            Exit Function
          End If
        End If
        If StatusError() <> "" Then
          bmsg = "SAP message: " & StatusError()
          Exit Function
        End If
        WScript.Sleep 500
      Loop
      If Not Find(ZAHLS) Is Nothing Then
        Trace "reached the Payment transactions screen with way " & way
        Exit For
      End If
      Trace "way " & way & " did not reach it (window title: " & TitleText() & ", status bar: " & StatusText() & ")"
    Next
    If Find(ZAHLS) Is Nothing Then
      bmsg = "the Payment block field did not appear after 3 ways (window title: " & TitleText() & ")"
      Exit Function
    End If
  End If

  cur = ""
  On Error Resume Next
  cur = UCase(Trim(session.findById(ZAHLS).text))
  If Err.Number <> 0 Then
    bmsg = "could not read the payment block: " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0
  Trace "current payment block: '" & cur & "', wanted: '" & desired & "'"

  If cur = desired Then
    bmsg = "already " & Describe(cur) & ", nothing to do"
    HandleBlock = True
    Exit Function
  End If
  If desired = "" Then
    If cur = "A" Then
      newVal = ""
    Else
      bmsg = "has " & Describe(cur) & ", left unchanged (only block A is removed)"
      HandleBlock = True
      Exit Function
    End If
  Else
    If cur = "" Then
      newVal = desired
    Else
      bmsg = "has " & Describe(cur) & ", left unchanged (wanted " & desired & ")"
      HandleBlock = True
      Exit Function
    End If
  End If

  If DRY_RUN Then
    bmsg = "dry run: would change " & Describe(cur) & " to " & Describe(newVal)
    HandleBlock = True
    Exit Function
  End If

  If Not SetText(ZAHLS, newVal) Then bmsg = gErr: Exit Function
  Trace "payment block '" & cur & "' -> '" & newVal & "', saving"
  If Not Press("wnd[0]/tbar[0]/btn[11]") Then bmsg = gErr: Exit Function

  ' after a successful Save SAP leaves the screen and shows a message
  t = Timer
  Do While Timer - t < 60
    If Not Find("wnd[1]") Is Nothing Then
      bmsg = "popup after Save: " & PopupTitle(1)
      Exit Function
    End If
    st = StatusType()
    If st = "E" Or st = "A" Then
      bmsg = "SAP error on Save: " & StatusText()
      Exit Function
    End If
    If Find(ZAHLS) Is Nothing Then Exit Do
    WScript.Sleep 500
  Loop
  If Not Find(ZAHLS) Is Nothing Then
    bmsg = "the screen did not close after Save (status bar: " & StatusText() & ")"
    Exit Function
  End If
  stText = StatusText()
  st = StatusType()
  If st = "E" Or st = "A" Then
    bmsg = "SAP error on Save: " & stText
    Exit Function
  End If
  If InStr(LCase(stText), "no change") > 0 Then
    bmsg = "SAP reports that nothing was changed: " & stText
    Exit Function
  End If
  bmsg = "changed " & Describe(cur) & " to " & Describe(newVal) & " (SAP: " & stText & ")"
  HandleBlock = True
End Function

' Returns True when the file was attached (or, in dry run, when the attachment list opened).
Function DoVendor(lifnr, pba, cc, ByRef msg)
  Dim grid, before, after, t, ready, i, title, ok3, errText, initScreen, dumped, lastLogged, picked, viaList, k, g2, blockMsg, blockOk
  DoVendor = False
  msg = ""
  If Find("wnd[0]") Is Nothing Then WScript.Sleep 5000
  If Find("wnd[0]") Is Nothing Then msg = "SAP session is no longer available (logged off?)": Exit Function
  ClosePopups

  Trace "opening XK02 for vendor " & lifnr
  ' 1. XK02 start screen, vendor number, Address view (as recorded)
  If Not SetText("wnd[0]/tbar[0]/okcd", "/nXK02") Then msg = gErr: Exit Function
  If Not SendKey("wnd[0]", 0) Then msg = gErr: Exit Function
  If WaitFor("wnd[0]/usr/ctxtRF02K-LIFNR", 60) Is Nothing Then msg = "XK02 start screen did not appear": Exit Function
  initScreen = ScreenNo()
  If Not SetText("wnd[0]/usr/ctxtRF02K-LIFNR", lifnr) Then msg = gErr: Exit Function
  ' the views block only shows after F8 (as in the recording) when the Payment transactions checkbox is missing
  If Find("wnd[0]/usr/chkRF02K-D0215") Is Nothing Then
    Trace "showing the view selection (F8)"
    If Not SendKey("wnd[0]", 8) Then msg = gErr: Exit Function
    If WaitFor("wnd[0]/usr/chkRF02K-D0215", 20) Is Nothing Then msg = "the company code Payment transactions checkbox did not appear": Exit Function
  End If
  If Not SetChecked("wnd[0]/usr/chkRF02K-D0110", True) Then msg = gErr: Exit Function
  If Not SetChecked("wnd[0]/usr/chkRF02K-D0215", True) Then msg = gErr: Exit Function
  If Not SetText("wnd[0]/usr/ctxtRF02K-BUKRS", cc) Then msg = gErr: Exit Function
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

  If ATTACH_EMAIL Then
  ' 3. Services for Object > Create attachment. (The "Attachment list" entry only exists for vendors that
  '    already have attachments, so the Create entry is used directly.)
  Trace "vendor screen is open (" & TitleText() & ", " & Int(Timer - t) & " s), opening Services for Object > Create attachment"
  ' Optional pause (LOAD_SECS), then select Services for Object > Create attachment ONCE. SAP can be very slow
  ' to react: it may report an error to the script while it is in fact still starting the dialog. So after
  ' every attempt the script only waits for the Import file dialog and presses the menu again only if
  ' nothing has appeared after DIALOG_SECS.
  If LOAD_SECS > 0 Then
    Trace "waiting " & LOAD_SECS & " s for SAP to finish loading the vendor"
    t = Timer
    Do While Timer - t < LOAD_SECS
      If Not Find("wnd[1]") Is Nothing Then ClosePopup 1
      WScript.Sleep 1000
    Loop
  End If

  ok3 = False
  For i = 1 To 2
    errText = ""
    picked = False
    viaList = False
    ' SAP shows an "Attachment list" popup by itself for some vendors: close it (never confirms anything)
    If Not Find("wnd[1]") Is Nothing Then
      If InStr(LCase(PopupTitle(1)), "attachment list") > 0 Then ClosePopup 1
    End If

    ' 1. System > Services for Object (a normal menu click) shows the toolbox toolbar
    If Find(TOOLBOX) Is Nothing Then
      Trace "attempt " & i & ": opening System > Services for Object"
      On Error Resume Next
      session.findById("wnd[0]/mbar/menu[5]/menu[6]").select
      If Err.Number <> 0 Then
        errText = "menu System > Services for Object: " & Err.Description
        Err.Clear
      End If
      On Error GoTo 0
      If errText = "" Then
        If WaitFor(TOOLBOX, 40) Is Nothing Then errText = "the Services for Object toolbar did not appear"
      End If
    End If

    If errText = "" Then
      BringSapToFront

      ' 2a. Route A: the "Attachment list" button (a plain button), then the list's own Create menu
      Trace "attempt " & i & ": toolbar is open, trying the Attachment list button"
      On Error Resume Next
      session.findById(TOOLBOX).pressButton "VIEW_ATTA"
      If Err.Number <> 0 Then
        Trace "Attachment list button: " & Err.Description
        Err.Clear
      End If
      On Error GoTo 0
      Set grid = WaitFor(LIST_SHELL, 40)
      If Not grid Is Nothing Then
        Trace "attachment list is open, opening its Create menu"
        before = RowCountOf(grid)
        BringSapToFront
        On Error Resume Next
        grid.pressToolbarContextButton "%ATTA_CREATE"
        If Err.Number <> 0 Then
          Trace "list Create button: " & Err.Description
          Err.Clear
        End If
        On Error GoTo 0
        t = Timer
        Do While Timer - t < MENU_SECS
          If Not Find("wnd[1]/usr/ctxtDY_PATH") Is Nothing Then
            picked = True
            Exit Do
          End If
          On Error Resume Next
          grid.selectContextMenuItem "%GOS_PCATTA_CREA"
          If Err.Number <> 0 Then
            Err.Clear
          Else
            picked = True
          End If
          On Error GoTo 0
          If picked Then Exit Do
          WScript.Sleep 2000
        Loop
        If picked Then
          viaList = True
          Trace "Create attachment selected from the list after " & Int(Timer - t) & " s (" & before & " attachments so far)"
        Else
          Trace "the list's Create menu was not ready after " & Int(Timer - t) & " s"
          If Find("wnd[1]/usr/ctxtDY_PATH") Is Nothing Then ClosePopup 1
        End If
      Else
        Trace "no attachment list appeared (the vendor may have no attachments yet)"
      End If

      ' 2b. Route B: the toolbar's own Create menu, opened ONCE
      If Not picked Then
        Trace "attempt " & i & ": trying the toolbar's Create menu"
        BringSapToFront
        On Error Resume Next
        session.findById(TOOLBOX).pressContextButton "CREATE_ATTA"
        If Err.Number <> 0 Then
          errText = "Create button: " & Err.Description
          Err.Clear
        End If
        On Error GoTo 0
        If errText = "" Then
          t = Timer
          Do While Timer - t < MENU_SECS
            If Not Find("wnd[1]/usr/ctxtDY_PATH") Is Nothing Then
              picked = True
              Exit Do
            End If
            On Error Resume Next
            session.findById(TOOLBOX).selectContextMenuItem "%GOS_PCATTA_CREA"
            If Err.Number <> 0 Then
              Err.Clear
              session.findById(TOOLBOX).selectContextMenuItemByText "Create attachment"
              If Err.Number <> 0 Then
                Err.Clear
              Else
                picked = True
              End If
            Else
              picked = True
            End If
            On Error GoTo 0
            If picked Then Exit Do
            WScript.Sleep 2000
          Loop
          If picked Then
            Trace "Create attachment selected from the toolbar after " & Int(Timer - t) & " s"
          Else
            errText = "the Create menu was not ready after " & Int(Timer - t) & " s"
          End If
        End If
      End If
    End If

    ' 3. SAP's own Import file dialog
    If picked Then
      errText = ""
      t = Timer
      If WaitForImportDialog(DIALOG_SECS) Then
        ok3 = True
        Trace "Import file dialog opened after " & Int(Timer - t) & " s"
        Exit For
      End If
      errText = "no Import file dialog"
    End If
    Trace "attempt " & i & " failed: " & errText
  Next
  If Not ok3 Then msg = "The Import file dialog did not appear (" & errText & ")": Exit Function

  If DRY_RUN Then
    msg = "dry run: the Import file dialog opened (nothing attached)"
    DoVendor = True
  Else
    ' 4. SAP's own "Import file" dialog: Directory + File Name, then OK
    Trace "filling the Import file dialog"
    If Not SetText("wnd[1]/usr/ctxtDY_PATH", ATTACH_FOLDER) Then msg = gErr: Exit Function
    If Not SetText("wnd[1]/usr/ctxtDY_FILENAME", ATTACH_NAME) Then msg = gErr: Exit Function
    If Not Press("wnd[1]/tbar[0]/btn[0]") Then msg = gErr: Exit Function

    ' 5. confirm the upload.
    '    Via the attachment list the list is re-opened after the upload and the rows are counted again
    '    (the status bar stays empty in this route). Via the toolbar's Create menu SAP shows
    '    "The attachment was successfully created".
    ready = False
    after = -1
    If viaList Then
      t = Timer
      Do While Timer - t < DIALOG_SECS
        If Not Find("wnd[2]") Is Nothing Then msg = "Import file error popup: " & PopupTitle(2): Exit Function
        If Find("wnd[1]/usr/ctxtDY_PATH") Is Nothing Then Exit Do
        WScript.Sleep 500
      Loop
      If before < 0 Then
        ready = True
        Trace "the attachment count could not be read before the upload, so the result cannot be verified"
      Else
        For k = 1 To 4
          WScript.Sleep 4000
          ClosePopups
          Trace "re-opening the attachment list to count again (try " & k & ")"
          Set g2 = OpenList()
          If Not g2 Is Nothing Then
            after = RowCountOf(g2)
            If after > before Then
              ready = True
              Exit For
            End If
          End If
        Next
      End If
      If Not ready Then msg = "the attachment was not confirmed: the list shows " & after & " attachments (was " & before & ")": Exit Function
    Else
      t = Timer
      Do While Timer - t < DIALOG_SECS
        If Not Find("wnd[2]") Is Nothing Then msg = "Import file error popup: " & PopupTitle(2): Exit Function
        If Find("wnd[1]/usr/ctxtDY_PATH") Is Nothing Then
          If InStr(LCase(StatusText()), "attachment") > 0 And InStr(LCase(StatusText()), "created") > 0 Then
            ready = True
            Exit Do
          End If
          If StatusError() <> "" Then msg = "SAP message: " & StatusError(): Exit Function
          If Not Find("wnd[1]") Is Nothing Then ClosePopup 1
        End If
        WScript.Sleep 500
      Loop
      If Not ready Then msg = "No 'attachment created' confirmation within " & DIALOG_SECS & " s (status bar: " & StatusText() & ")": Exit Function
    End If
    If viaList Then
      msg = "attached via the attachment list (" & before & " -> " & after & " attachments)"
    Else
      msg = "attached (" & StatusText() & ")"
    End If
  End If
  Trace "closing the attachment list / dialogs (if open)"
  ClosePopups
  End If

  ' payment block (rules above)
  blockMsg = ""
  blockOk = HandleBlock(pba, blockMsg)
  If msg <> "" Then msg = msg & " | "
  If blockOk Then
    msg = msg & "payment block: " & blockMsg
  Else
    msg = msg & "PAYMENT BLOCK PROBLEM: " & blockMsg
  End If
  DoVendor = blockOk

  ' 6. dismiss any dialog and leave XK02 (without saving anything that was not saved above)
  Trace "leaving XK02"
  ClosePopups
  If SetText("wnd[0]/tbar[0]/okcd", "/n") Then SendKey "wnd[0]", 0
  WScript.Sleep 1500
  ClosePopups
End Function

' ---- only one copy may run: several copies pressing the same SAP window at once make everything fail ----
If InstancesRunning() > 1 Then
  Log "STOPPED: another copy of this script is already running."
  MsgBox "Another copy of this script is already running." & vbCrLf & "Press Ctrl+Shift+Esc, end every wscript.exe under Details/Processes, then start it again.", 48, "Payment block"
  WScript.Quit 1
End If

' ---- attach to the running SAP session ----
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
  Log "ERROR: SAP GUI is not running."
  MsgBox "SAP GUI is not running. Open SAP and log in first.", 16, "Payment block"
  WScript.Quit 1
End If
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children(0)
Set session = connection.Children(0)
If Err.Number <> 0 Then
  Log "ERROR: could not attach to a logged-in SAP session."
  MsgBox "Could not attach to a logged-in SAP session.", 16, "Payment block"
  WScript.Quit 1
End If
On Error GoTo 0
session.findById("wnd[0]").maximize

Dim vendors, k, done, okCount, failCount, fails, msg, started, stoppedEarly, doneDict, doneKey, skipped, tsf, parts, vp, failList, noteList
done = 0: okCount = 0: failCount = 0: fails = 0: stoppedEarly = False: failList = "": noteList = ""
If Not fso.FileExists(VENDOR_FILE) Then
  MsgBox "Vendor file not found: " & VENDOR_FILE, 16, "Payment block"
  WScript.Quit 1
End If
Set doneDict = CreateObject("Scripting.Dictionary")
skipped = 0
If SKIP_DONE And fso.FileExists(doneFile) Then
  Set tsf = fso.OpenTextFile(doneFile, 1)
  Do While Not tsf.AtEndOfStream
    parts = Split(tsf.ReadLine, "|")
    If UBound(parts) >= 3 Then doneDict(parts(0) & "|" & parts(1) & "|" & parts(2) & "|" & parts(3)) = True
  Loop
  tsf.Close
End If
Set vendors = ReadVendors()
If DRY_RUN Then runLabel = "DRY RUN" Else runLabel = "REAL RUN"
Log "===== " & runLabel & ": " & vendors.Count & " vendors in " & VENDOR_FILE & " | file " & ATTACH_FOLDER & "\" & ATTACH_NAME & " ====="

For Each k In vendors.Keys
  If k >= START_ROW Then
    vp = Split(vendors(k), "|")
    doneKey = vp(0) & "|" & vp(2) & "|" & vp(1) & "|" & LCase(ATTACH_NAME)
    If SKIP_DONE And doneDict.Exists(doneKey) Then
      skipped = skipped + 1
      Log "row " & k & "  vendor " & vp(0) & "  SKIPPED  already done in an earlier run (payment-block-done.txt)"
    Else
      If MAX_VENDORS > 0 And done >= MAX_VENDORS Then Exit For
      done = done + 1
      started = Timer
      msg = ""
      If DoVendor(vp(0), vp(1), vp(2), msg) Then
        okCount = okCount + 1
        fails = 0
        Log "row " & k & "  vendor " & vp(0) & "  OK   " & msg & "  (" & Int(Timer - started) & " s)"
        If InStr(msg, "left unchanged") > 0 Then noteList = noteList & vbCrLf & "row " & k & "  vendor " & vp(0) & ": " & msg
        If Not DRY_RUN Then
          Set tsf = fso.OpenTextFile(doneFile, 8, True)
          tsf.WriteLine doneKey & "|" & Now
          tsf.Close
          doneDict(doneKey) = True
        End If
      Else
        failCount = failCount + 1
        fails = fails + 1
        Log "row " & k & "  vendor " & vp(0) & "  FAILED   " & msg
        failList = failList & vbCrLf & "row " & k & "  vendor " & vp(0) & ": " & msg
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
  End If
Next

Log "===== finished: " & okCount & " ok, " & failCount & " failed, " & skipped & " skipped" & IIf2(stoppedEarly) & " ====="
Dim summary, icon
summary = runLabel & " finished." & vbCrLf & okCount & " ok, " & failCount & " failed, " & skipped & " skipped" & IIf2(stoppedEarly) & "."
icon = 64
If failCount > 0 Then
  summary = summary & vbCrLf & vbCrLf & "PROBLEMS - please check:" & failList
  icon = 16
End If
If noteList <> "" Then
  summary = summary & vbCrLf & vbCrLf & "Left unchanged on purpose (another block is in place):" & noteList
  If icon = 64 Then icon = 48
End If
summary = summary & vbCrLf & vbCrLf & "Details: " & logPath
MsgBox summary, icon, "Payment block"

Function IIf2(flag)
  If flag Then IIf2 = " (stopped early)" Else IIf2 = ""
End Function
