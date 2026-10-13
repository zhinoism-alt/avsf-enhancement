' Excel (first sheet, row 1 = header): column A = vendor number. Optional column with the header "CC" = company code of the
' vendor (SAP then also checks that the vendor exists in that company code; blank = no company code). Optional column with the header "File" = the full path of the
' file to attach to that vendor (blank or no such column = the default file in the settings below).
' Attaches a file to a list of vendors in SAP: the vendor is opened in display mode (XK03, nothing can be changed and
' the vendor is not locked), then Services for Object > Create attachment. Set TCODE = "XK02" to use change mode instead.
' Built from a recording. SAP GUI must be open and logged in. It uses your first SAP window,
' so leave SAP alone while it runs. Results are written to attach-log.txt next to this script.
Option Explicit

' ===== Settings =====
Dim DEFAULT_FOLDER, DEFAULT_NAME
Dim LOG_FOLDER
Dim VENDOR_FILE, ATTACH_FOLDER, ATTACH_NAME, DRY_RUN, CHECK_ONLY, TCODE, DEFAULT_CC, START_ROW, MAX_VENDORS, MAX_FAILS_IN_ROW, WAIT_SECS, LOAD_SECS, DIALOG_SECS, MENU_SECS, SKIP_DONE
VENDOR_FILE = "Attachment Script.xlsx"   ' Excel list: searched in the "lists" folder, then next to the script; or give a full path
ATTACH_FOLDER = ""   ' default attachment folder ("" = the "attachments" folder, or the base folder if there is none)
ATTACH_NAME = "RE_ Suspensiones hasta nuevo aviso.msg"
LOG_FOLDER = "logs"      ' folder for the log and the done list (relative to the base folder; created when missing; "" = the folder of this script)
DRY_RUN = False           ' True = open each vendor and its attachment list but attach NOTHING
DEFAULT_CC = ""           ' company code used when the CC column is blank or missing ("" = open the vendor without a company code)
TCODE = "XK03"           ' transaction used to open each vendor: XK03 = display (recommended), XK02 = change
CHECK_ONLY = False        ' True = only run the start-up check (SAP, Excel, files) and show a report; touches nothing in SAP
START_ROW = 2             ' first Excel row to process (row 1 is the header)
MAX_VENDORS = 0           ' 0 = every vendor in the Excel file; use a small number to test
MAX_FAILS_IN_ROW = 3      ' stop after this many failures in a row
SKIP_DONE = True          ' skip a vendor that already got THIS file TODAY (protects against double runs; a new day always runs again; list in attach-done.txt)
WAIT_SECS = 120           ' how long to wait for slow SAP screens
LOAD_SECS = 0             ' optional pause after the vendor opens before touching Services for Object (0 = none)
MENU_SECS = 45            ' how long SAP may take to fill the Services for Object menu after the button is pressed
DIALOG_SECS = 150         ' how long SAP may take to open the Import file dialog after Create attachment is selected
' ====================

' tidy the settings: remove stray quote marks (e.g. from "Copy as path") and trailing backslashes
VENDOR_FILE = FindListFile(VENDOR_FILE)
ATTACH_FOLDER = DefaultAttachFolder(ATTACH_FOLDER)
ATTACH_NAME = CleanPath(ATTACH_NAME)
DEFAULT_FOLDER = ATTACH_FOLDER
DEFAULT_NAME = ATTACH_NAME

Function CleanPath(p)
  p = Trim(Replace(p, Chr(34), ""))
  Do While Right(p, 1) = "\"
    p = Left(p, Len(p) - 1)
  Loop
  CleanPath = p
End Function

Function ScriptDir()
  ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
End Function

' ---- folder layout: scripts\ lists\ attachments\ logs\ (the old "everything in one folder" layout still works) ----
Function FsoGet()
  Set FsoGet = CreateObject("Scripting.FileSystemObject")
End Function

Function IsAbsolutePath(p)
  IsAbsolutePath = (InStr(p, ":") > 0 Or Left(p, 2) = "\\")
End Function

' base folder: the parent of the folder this script is in when that folder is called "scripts", else the script folder itself
Function BaseDir()
  Dim sd
  sd = ScriptDir()
  If LCase(Mid(sd, InStrRev(sd, "\") + 1)) = "scripts" Then
    BaseDir = FsoGet().GetParentFolderName(sd)
  Else
    BaseDir = sd
  End If
End Function

' the vendor list: a full path is used as it is; otherwise it is searched in <base>\lists, <base> and the script folder
Function FindListFile(p)
  Dim f, c, i
  p = CleanPath(p)
  If IsAbsolutePath(p) Then
    FindListFile = p
    Exit Function
  End If
  Set f = FsoGet()
  c = Array(BaseDir() & "\lists\" & p, BaseDir() & "\" & p, ScriptDir() & "\" & p)
  For i = 0 To UBound(c)
    If f.FileExists(c(i)) Then
      FindListFile = c(i)
      Exit Function
    End If
  Next
  FindListFile = c(0)
End Function

' default folder of the attachments: <base>\attachments when it exists, else the base folder
Function DefaultAttachFolder(p)
  p = CleanPath(p)
  If p = "" Then
    If FsoGet().FolderExists(BaseDir() & "\attachments") Then
      DefaultAttachFolder = BaseDir() & "\attachments"
    Else
      DefaultAttachFolder = BaseDir()
    End If
  ElseIf IsAbsolutePath(p) Then
    DefaultAttachFolder = p
  Else
    DefaultAttachFolder = BaseDir() & "\" & p
  End If
End Function

' folder for the log and the done list (created when missing; falls back to the script folder)
Function ResolveLogDir(p)
  Dim d, f
  p = CleanPath(p)
  Set f = FsoGet()
  If p = "" Then
    ResolveLogDir = ScriptDir()
    Exit Function
  End If
  If IsAbsolutePath(p) Then
    d = p
  Else
    d = BaseDir() & "\" & p
  End If
  On Error Resume Next
  If Not f.FolderExists(d) Then f.CreateFolder d
  If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
  If f.FolderExists(d) Then
    ResolveLogDir = d
  Else
    ResolveLogDir = ScriptDir()
  End If
End Function

Function NeedsFile()
  NeedsFile = True
End Function

Const LIST_SHELL = "wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell"
Const TOOLBOX = "wnd[0]/shellcont/shell"   ' toolbar with CREATE_ATTA, VIEW_ATTA, ... shown by System > Services for Object

Dim fso, session, application, connection, SapGuiAuto, logPath, doneFile, gErr, runLabel, gSapInfo, logDir, gBefore, gAfter, resultsFile, historyFile, resultLines
Set fso = CreateObject("Scripting.FileSystemObject")
logDir = ResolveLogDir(LOG_FOLDER)
logPath = fso.BuildPath(logDir, "attach-log.txt")
doneFile = fso.BuildPath(logDir, "attach-done.txt")
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

' XK02/XK03 remember the last view selection AND the company code, so the start screen is reset every time:
' ONLY Address (general data, D0110) is ticked; company code and purchasing org are cleared.
Sub SetViewChecks(node)
  Dim i, n, c
  On Error Resume Next
  If node.Type = "GuiCheckBox" Then
    node.selected = (InStr(node.Id, "RF02K-D0110") > 0)
    Err.Clear
  Else
    n = node.Children.Count
    If Err.Number = 0 Then
      For i = 0 To n - 1
        Set c = node.Children(CLng(i))
        SetViewChecks c
      Next
    End If
    Err.Clear
  End If
  On Error GoTo 0
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

' ---- start-up check: tells the user in plain words what to fix before anything touches SAP ----
Function Preflight()
  Dim problems, xl, tf, usr, lng, sysName, clientNo, cnt, wtxt, bad
  problems = ""
  gSapInfo = ""

  ' 1. the script folder must be writable (log and done list are written next to the script)
  On Error Resume Next
  Set tf = fso.OpenTextFile(logPath, 8, True)
  If Err.Number <> 0 Then
    problems = problems & vbCrLf & "- The folder of this script is read-only. Unzip the files into a normal folder (for example C:\SAP-scripts) and run the script from there. (Log folder: " & logDir & ")"
    Err.Clear
  Else
    tf.Close
  End If

  ' 2. Excel is needed to read the vendor list
  Set xl = Nothing
  Set xl = CreateObject("Excel.Application")
  If Err.Number <> 0 Then
    problems = problems & vbCrLf & "- Excel is not available on this PC (the script reads the vendor list through Excel)."
    Err.Clear
  Else
    xl.Quit
  End If
  Set xl = Nothing
  On Error GoTo 0

  ' 3. the vendor list
  If Not fso.FileExists(VENDOR_FILE) Then
    problems = problems & vbCrLf & "- The vendor Excel file was not found: " & VENDOR_FILE & vbCrLf & "  Put it in the ""lists"" folder (or next to the script) with that name, or change VENDOR_FILE at the top of the script."
  End If

  ' 4. SAP GUI, scripting, logon, language
  On Error Resume Next
  Set SapGuiAuto = Nothing
  Set SapGuiAuto = GetObject("SAPGUI")
  If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
    Err.Clear
    problems = problems & vbCrLf & "- SAP GUI is not running. Open SAP Logon, log on, then start the script."
  Else
    Set application = Nothing
    Set application = SapGuiAuto.GetScriptingEngine
    bad = False
    If Err.Number <> 0 Then
      bad = True
      Err.Clear
    End If
    If application Is Nothing Then bad = True
    If bad Then
      problems = problems & vbCrLf & "- SAP GUI scripting is switched off on this PC. In SAP press Alt+F12, choose Options > Accessibility & Scripting > Scripting, tick 'Enable scripting' and untick the two 'Notify when...' boxes. If the option is greyed out, ask IT."
    Else
      cnt = 0
      cnt = application.Children.Count
      Err.Clear
      If cnt = 0 Then
        problems = problems & vbCrLf & "- No SAP system is open. Log on to SAP first."
      Else
        ' look at every connection and session, prefer one that is logged on
        Dim ci, si, nConn, nSess, cand, firstSess, sessErr
        Set session = Nothing
        Set firstSess = Nothing
        nConn = application.Children.Count
        nSess = 0
        For ci = 0 To nConn - 1
          Set connection = Nothing
          Set connection = application.Children(CLng(ci))
          If Err.Number = 0 Then
            cnt = 0
            cnt = connection.Children.Count
            For si = 0 To cnt - 1
              Set cand = Nothing
              Set cand = connection.Children(CLng(si))
              If Err.Number = 0 Then
                nSess = nSess + 1
                If firstSess Is Nothing Then Set firstSess = cand
                If session Is Nothing Then
                  If cand.info.user <> "" Then Set session = cand
                End If
              End If
              Err.Clear
            Next
          End If
          Err.Clear
        Next
        If session Is Nothing Then Set session = firstSess
        If session Is Nothing Then
          ' fallback: whatever SAP says is the active session
          Set cand = Nothing
          Set cand = application.ActiveSession
          If Err.Number = 0 Then
            If Not cand Is Nothing Then
              Set session = cand
              nSess = nSess + 1
            End If
          Else
            sessErr = Err.Number & " " & Err.Description
          End If
          Err.Clear
        End If
        cnt = nSess
        Err.Clear
        If cnt = 0 Then
          problems = problems & vbCrLf & "- No SAP session is open (SAP sees " & nConn & " connection(s) and 0 sessions; ActiveSession: " & sessErr & "). Log on to SAP first and wait until the SAP Easy Access screen shows. If you are already logged on, close the SAP window, open SAP Logon again and log on again (a stale connection is the usual cause)."
        Else
          wtxt = ""
          wtxt = session.findById("wnd[0]").Text
          If Err.Number <> 0 Then
            Err.Clear
            problems = problems & vbCrLf & "- SAP does not accept scripts (scripting may be disabled for your user on the SAP server). Ask your SAP administrator."
          Else
            usr = ""
            usr = session.info.user
            lng = ""
            lng = UCase(session.info.language)
            sysName = session.info.systemName
            clientNo = session.info.client
            Err.Clear
            If usr = "" Then
              problems = problems & vbCrLf & "- You are not logged on in SAP (the logon screen is open). Log on first."
            ElseIf lng <> "EN" Then
              problems = problems & vbCrLf & "- SAP is logged on in language " & lng & ". The script recognises English screens only. Log off and log on again with language EN."
            Else
              gSapInfo = "SAP user " & usr & ", system " & sysName & " client " & clientNo & ", language " & lng
            End If
          End If
        End If
      End If
    End If
  End If
  On Error GoTo 0
  Preflight = problems
End Function

' report for CHECK_ONLY = True: SAP and Excel are fine, how many vendors, which attachment files are missing
Sub CheckOnlyReport()
  Dim vend, kk, vpp, full, missing, total, txt
  Set vend = ReadVendors()
  total = vend.Count
  missing = ""
  If NeedsFile() Then
    For Each kk In vend.Keys
      vpp = Split(vend(kk), "|")
      full = ResolveFile(vpp(UBound(vpp)))
      If Not fso.FileExists(full) Then missing = missing & vbCrLf & "  row " & kk & ": " & full
    Next
  End If
  txt = "All checks passed." & vbCrLf & vbCrLf & gSapInfo & vbCrLf & "Attachments folder: " & DEFAULT_FOLDER & vbCrLf & "Logs folder: " & logDir & vbCrLf & "Vendor list: " & VENDOR_FILE & " (" & total & " vendors)"
  If missing <> "" Then
    txt = txt & vbCrLf & vbCrLf & "ATTACHMENT FILES NOT FOUND:" & missing & vbCrLf & "(put the files there, or fix the File column)"
    MsgBox txt, 48, "Attach document"
  Else
    txt = txt & vbCrLf & vbCrLf & "Nothing was changed in SAP. Set CHECK_ONLY = False to run the script."
    MsgBox txt, 64, "Attach document"
  End If
  If missing <> "" Then
    Log "Check only: " & gSapInfo & "; " & total & " vendors; ATTACHMENT FILES MISSING:" & missing
  Else
    Log "Check only: " & gSapInfo & "; " & total & " vendors; all attachment files found"
  End If
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

' the file for one vendor: a full path from the Excel "File" column; a bare file name is looked up in the default folder;
' blank = the default file from the settings
Function ResolveFile(p)
  Dim f, c, i
  p = CleanPath(p)
  If p = "" Then
    ResolveFile = DEFAULT_FOLDER & "\" & DEFAULT_NAME
    Exit Function
  End If
  If IsAbsolutePath(p) Then
    ResolveFile = p
    Exit Function
  End If
  Set f = FsoGet()
  c = Array(DEFAULT_FOLDER & "\" & p, BaseDir() & "\attachments\" & p, BaseDir() & "\" & p, ScriptDir() & "\" & p)
  For i = 0 To UBound(c)
    If f.FileExists(c(i)) Then
      ResolveFile = c(i)
      Exit Function
    End If
  Next
  ResolveFile = c(0)
End Function

Function FindCCColumn(ws)
  Dim c, lastCol, v
  FindCCColumn = 0
  lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column
  For c = 1 To lastCol
    v = LCase(Trim(CStr(ws.Cells(1, c).Value)))
    If v = "cc" Or v = "company code" Or v = "company" Or v = "bukrs" Or v = "co code" Then FindCCColumn = c
  Next
End Function

Function FindFileColumn(ws)
  Dim c, lastCol, v
  FindFileColumn = 0
  lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column
  For c = 1 To lastCol
    v = LCase(Trim(CStr(ws.Cells(1, c).Value)))
    If v = "file" Or v = "file path" Or v = "attachment" Or v = "path" Then FindFileColumn = c
  Next
End Function

Function IIfCC(cc)
  If cc = "" Then
    IIfCC = ""
  Else
    IIfCC = " (CC " & cc & ")"
  End If
End Function

Function Pad(v, n)
  Dim t
  t = CStr(v)
  If Len(t) >= n Then
    Pad = Left(t, n - 1) & " "
  Else
    Pad = t & Space(n - Len(t))
  End If
End Function

Function Q(v)
  Q = Chr(34) & Replace(CStr(v), Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Sub WriteCsvLine(path, line)
  Dim f
  On Error Resume Next
  Set f = fso.OpenTextFile(path, 8, True, -1)
  If Err.Number = 0 Then
    f.WriteLine line
    f.Close
  Else
    Err.Clear
  End If
  On Error GoTo 0
End Sub

' one line per Excel row in the results list (a CSV that opens in Excel) and in the log
Sub AddResult(rowNo, lifnr, cc, fullFile, status, bef, aft, note, secs)
  Dim counts, line
  counts = ""
  If bef >= 0 And aft >= 0 Then counts = bef & " -> " & aft
  line = Q(Now) & "," & Q(rowNo) & "," & Q(lifnr) & "," & Q(cc) & "," & Q(fso.GetFileName(fullFile)) & "," & Q(fullFile) & "," & Q(status) & "," & Q(counts) & "," & Q(note) & "," & Q(secs)
  WriteCsvLine resultsFile, line
  WriteCsvLine historyFile, line
  resultLines = resultLines & vbCrLf & Pad(status, 10) & Pad("row " & rowNo, 8) & Pad(lifnr, 12) & Pad("CC " & cc, 9) & Pad(counts, 9) & fso.GetFileName(fullFile)
End Sub

Function TodayStamp()
  TodayStamp = Year(Date) & "-" & Right("0" & Month(Date), 2) & "-" & Right("0" & Day(Date), 2)
End Function

Function IsVendorNumber(s)
  Dim re
  Set re = CreateObject("VBScript.RegExp")
  re.Pattern = "^[0-9]{1,10}$"
  IsVendorNumber = re.Test(s)
End Function

Function ReadVendors()
  Dim xl, wb, ws, lastRow, r, v, d, fileCol, fileVal, ccCol, ccVal
  Set d = CreateObject("Scripting.Dictionary")
  Set xl = CreateObject("Excel.Application")
  xl.Visible = False
  xl.DisplayAlerts = False
  Set wb = xl.Workbooks.Open(VENDOR_FILE, 0, True)
  Set ws = wb.Worksheets(1)
  lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
  fileCol = FindFileColumn(ws)
  ccCol = FindCCColumn(ws)
  For r = 1 To lastRow
    v = Trim(CStr(ws.Cells(r, 1).Value))
    If IsVendorNumber(v) Then
      fileVal = ""
      If fileCol > 0 Then fileVal = Trim(CStr(ws.Cells(r, fileCol).Value))
      ccVal = DEFAULT_CC
      If ccCol > 0 Then
        If Trim(CStr(ws.Cells(r, ccCol).Value)) <> "" Then ccVal = Trim(CStr(ws.Cells(r, ccCol).Value))
      End If
      d.Add r, Right("0000000000" & v, 10) & "|" & ccVal & "|" & fileVal
    End If
  Next
  wb.Close False
  xl.Quit
  Set ReadVendors = d
End Function

' System > Services for Object, found by its TEXT: in XK03 (display) the menu bar has no Edit menu, so the
' position of the System menu differs from XK02.
Function SelectServicesMenu()
  Dim bar, i, n, m, j, k, it
  SelectServicesMenu = False
  On Error Resume Next
  Set bar = Nothing
  Set bar = session.findById("wnd[0]/mbar")
  If Err.Number <> 0 Or bar Is Nothing Then
    gErr = "could not read the menu bar"
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  n = bar.Children.Count
  For i = 0 To n - 1
    Set m = bar.Children(CLng(i))
    If LCase(Trim(m.Text)) = "system" Then
      k = m.Children.Count
      For j = 0 To k - 1
        Set it = m.Children(CLng(j))
        If LCase(Trim(it.Text)) = "services for object" Then
          Err.Clear
          it.select
          If Err.Number <> 0 Then
            gErr = "menu System > Services for Object: " & Err.Description
            Err.Clear
          Else
            SelectServicesMenu = True
          End If
          On Error GoTo 0
          Exit Function
        End If
      Next
    End If
  Next
  gErr = "menu System > Services for Object was not found"
  Err.Clear
  On Error GoTo 0
End Function

' opens the Attachment list (System > Services for Object if the toolbox is not open yet); returns the grid or Nothing
Function OpenList()
  Dim g
  Set g = Nothing
  If Find(TOOLBOX) Is Nothing Then
    SelectServicesMenu
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

' Returns True when the file was attached (or, in dry run, when the attachment list opened).
Function DoVendor(lifnr, cc, ByRef msg)
  Dim grid, before, after, t, ready, i, title, ok3, errText, initScreen, dumped, lastLogged, picked, viaList, k, g2
  DoVendor = False
  msg = ""
  If Find("wnd[0]") Is Nothing Then WScript.Sleep 5000
  If Find("wnd[0]") Is Nothing Then msg = "SAP session is no longer available (logged off?)": Exit Function
  ClosePopups

  Trace "opening " & TCODE & " for vendor " & lifnr & IIfCC(cc)
  ' 1. start screen of the transaction, vendor number, Address view (as recorded)
  If Not SetText("wnd[0]/tbar[0]/okcd", "/n" & TCODE) Then msg = gErr: Exit Function
  If Not SendKey("wnd[0]", 0) Then msg = gErr: Exit Function
  If WaitFor("wnd[0]/usr/ctxtRF02K-LIFNR", 60) Is Nothing Then msg = TCODE & " start screen did not appear": Exit Function
  initScreen = ScreenNo()
  SetViewChecks session.findById("wnd[0]/usr")
  If Not SetChecked("wnd[0]/usr/chkRF02K-D0110", True) Then msg = gErr: Exit Function
  SetText "wnd[0]/usr/ctxtRF02K-BUKRS", cc
  SetText "wnd[0]/usr/ctxtRF02K-EKORG", ""
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
      If Not SelectServicesMenu() Then errText = gErr
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
        gBefore = before
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
    gAfter = after
    If viaList Then
      msg = "attached via the attachment list (" & before & " -> " & after & " attachments)"
    Else
      msg = "attached (" & StatusText() & ")"
    End If
    DoVendor = True
    Trace "closing the attachment list (if open)"
    ClosePopups
    If UCase(TCODE) = "XK02" Then
      ' change mode, as in the recording: press Save (nothing else changed, so SAP only reports "no changes")
      Press "wnd[0]/tbar[0]/btn[11]"
      WScript.Sleep 1500
      ClosePopups
    End If
  End If

  ' 6. dismiss any dialog and leave the transaction (attachments are stored as soon as they are created)
  Trace "leaving " & TCODE
  ClosePopups
  If SetText("wnd[0]/tbar[0]/okcd", "/n") Then SendKey "wnd[0]", 0
  WScript.Sleep 1500
  ClosePopups
End Function

' ---- only one copy may run: several copies pressing the same SAP window at once make everything fail ----
If InstancesRunning() > 1 Then
  Log "STOPPED: another copy of this script is already running."
  MsgBox "Another copy of this script is already running." & vbCrLf & "Press Ctrl+Shift+Esc, end every wscript.exe under Details/Processes, then start it again.", 48, "Attach document"
  WScript.Quit 1
End If

' ---- start-up check (also attaches to the SAP session) ----
Dim pfProblems
pfProblems = Preflight()
If pfProblems <> "" Then
  Log "START-UP CHECK FAILED:" & pfProblems
  MsgBox "The script cannot start yet. Please fix:" & vbCrLf & pfProblems, 48, "Attach document"
  WScript.Quit 1
End If
Log "Start-up check passed: " & gSapInfo
If CHECK_ONLY Then
  CheckOnlyReport
  WScript.Quit 0
End If
session.findById("wnd[0]").maximize

Dim vendors, k, done, okCount, failCount, fails, msg, started, stoppedEarly, doneDict, doneKey, skipped, tsf, parts, vp, fullPath, failList, occ, occKey, occN, processed, stampNow, histHeader, notRun
done = 0: okCount = 0: failCount = 0: fails = 0: stoppedEarly = False: failList = ""
If Not fso.FileExists(VENDOR_FILE) Then
  MsgBox "Vendor file not found: " & VENDOR_FILE, 16, "Attach document"
  WScript.Quit 1
End If
Set doneDict = CreateObject("Scripting.Dictionary")
skipped = 0
If SKIP_DONE And fso.FileExists(doneFile) Then
  Set tsf = fso.OpenTextFile(doneFile, 1)
  Do While Not tsf.AtEndOfStream
    parts = Split(tsf.ReadLine, "|")
    If UBound(parts) >= 4 Then doneDict(parts(0) & "|" & parts(1) & "|" & parts(2) & "|" & parts(3) & "|" & parts(4)) = True
  Loop
  tsf.Close
End If
Set vendors = ReadVendors()
If DRY_RUN Then runLabel = "DRY RUN" Else runLabel = "REAL RUN"
' results list: one CSV per run plus a running history of all runs (both open in Excel)
stampNow = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2) & "-" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
resultsFile = fso.BuildPath(logDir, "attach-results-" & stampNow & ".csv")
historyFile = fso.BuildPath(logDir, "attach-results-history.csv")
histHeader = "Time,Excel row,Vendor,Company code,File name,File path,Status,Attachments (before -> after),Message,Seconds"
WriteCsvLine resultsFile, histHeader
If Not fso.FileExists(historyFile) Then WriteCsvLine historyFile, histHeader
resultLines = ""
Set occ = CreateObject("Scripting.Dictionary")
Set processed = CreateObject("Scripting.Dictionary")
Log "===== " & runLabel & ": " & vendors.Count & " vendors in " & VENDOR_FILE & " | default file " & DEFAULT_FOLDER & "\" & DEFAULT_NAME & " ====="

For Each k In vendors.Keys
  vp = Split(vendors(k), "|")
  fullPath = ResolveFile(vp(2))
  ' the same vendor + company code + file may appear several times (several copies): each occurrence is its own job
  occKey = vp(0) & "|" & vp(1) & "|" & LCase(fullPath)
  occN = 1
  If occ.Exists(occKey) Then occN = occ(occKey) + 1
  occ(occKey) = occN
  If k >= START_ROW Then
    doneKey = occKey & "|" & occN & "|" & TodayStamp()
    If SKIP_DONE And doneDict.Exists(doneKey) Then
      skipped = skipped + 1
      processed(k) = True
      Log "row " & k & "  vendor " & vp(0) & "  SKIPPED  already done today with this file (copy " & occN & ")"
      AddResult k, vp(0), vp(1), fullPath, "SKIPPED", -1, -1, "already done today (copy " & occN & " of this vendor/company code/file)", 0
    ElseIf Not fso.FileExists(fullPath) Then
      failCount = failCount + 1
      processed(k) = True
      Log "row " & k & "  vendor " & vp(0) & "  FAILED   attachment file not found: " & fullPath
      failList = failList & vbCrLf & "row " & k & "  vendor " & vp(0) & ": attachment file not found: " & fullPath
      AddResult k, vp(0), vp(1), fullPath, "FAILED", -1, -1, "attachment file not found", 0
    Else
      If MAX_VENDORS > 0 And done >= MAX_VENDORS Then Exit For
      done = done + 1
      processed(k) = True
      started = Timer
      msg = ""
      gBefore = -1
      gAfter = -1
      ATTACH_FOLDER = fso.GetParentFolderName(fullPath)
      ATTACH_NAME = fso.GetFileName(fullPath)
      If DoVendor(vp(0), vp(1), msg) Then
        okCount = okCount + 1
        fails = 0
        Log "row " & k & "  vendor " & vp(0) & "  OK   " & IIfCC(vp(1)) & " [" & ATTACH_NAME & "]  " & msg & "  (" & Int(Timer - started) & " s)"
        If DRY_RUN Then
          AddResult k, vp(0), vp(1), fullPath, "DRY RUN OK", gBefore, gAfter, msg, Int(Timer - started)
        Else
          AddResult k, vp(0), vp(1), fullPath, "OK", gBefore, gAfter, msg, Int(Timer - started)
          Set tsf = fso.OpenTextFile(doneFile, 8, True)
          tsf.WriteLine doneKey & "|" & Now
          tsf.Close
          doneDict(doneKey) = True
        End If
      Else
        failCount = failCount + 1
        fails = fails + 1
        Log "row " & k & "  vendor " & vp(0) & "  FAILED   " & IIfCC(vp(1)) & " [" & ATTACH_NAME & "]  " & msg
        failList = failList & vbCrLf & "row " & k & "  vendor " & vp(0) & ": " & msg
        AddResult k, vp(0), vp(1), fullPath, "FAILED", gBefore, gAfter, msg, Int(Timer - started)
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

' rows that were not reached (stopped early or MAX_VENDORS) are listed too, so the list is complete
notRun = 0
For Each k In vendors.Keys
  If k >= START_ROW And Not processed.Exists(k) Then
    vp = Split(vendors(k), "|")
    notRun = notRun + 1
    AddResult k, vp(0), vp(1), ResolveFile(vp(2)), "NOT RUN", -1, -1, "not reached (run stopped early or MAX_VENDORS reached)", 0
  End If
Next

Log "RESULT LIST (" & runLabel & "):" & vbCrLf & Pad("STATUS", 10) & Pad("ROW", 8) & Pad("VENDOR", 12) & Pad("COMPANY", 9) & Pad("ATTACHED", 9) & "FILE" & resultLines
Log "===== finished: " & okCount & " ok, " & failCount & " failed, " & skipped & " skipped, " & notRun & " not run" & IIf2(stoppedEarly) & " ====="
Dim summary, icon
summary = runLabel & " finished." & vbCrLf & okCount & " ok, " & failCount & " failed, " & skipped & " skipped, " & notRun & " not run" & IIf2(stoppedEarly) & "."
icon = 64
If failCount > 0 Then
  summary = summary & vbCrLf & vbCrLf & "PROBLEMS - please check:" & failList
  icon = 16
End If
If skipped > 0 Then summary = summary & vbCrLf & vbCrLf & "Skipped = already done TODAY with the same settings. To repeat, delete logs\attach-done.txt" & " or set SKIP_DONE = False. (Several identical rows are fine: each row is its own copy.)"
summary = summary & vbCrLf & vbCrLf & "Result list (open in Excel): " & resultsFile & vbCrLf & "Log: " & logPath
MsgBox summary, icon, "Attach document"

Function IIf2(flag)
  If flag Then IIf2 = " (stopped early)" Else IIf2 = ""
End Function
