' EXPLORER: finds the toolbars on the current SAP screen, then clicks System > Services for Object ONCE
' (a normal menu click, nothing is pressed repeatedly) and records what appears. Writes gos-explore.txt.
' Start it on a vendor "Change Vendor: Address" screen with NO popups open. It does not attach anything.
Option Explicit

Dim fso, outFile, session, application, connection, SapGuiAuto
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "gos-explore.txt"), True, True)

Sub W(txt)
  outFile.WriteLine txt
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

Sub DescribeShell(o)
  Dim t, bc, b
  On Error Resume Next
  t = ""
  t = o.SubType
  W "   " & o.Id & "  |  " & o.Type & " / " & t
  bc = -1
  bc = o.ButtonCount
  If Err.Number = 0 And bc >= 0 Then
    For b = 0 To bc - 1
      W "      button " & b & ": id=" & o.GetButtonId(b) & "  text=" & o.GetButtonText(b) & "  tip=" & o.GetButtonTooltip(b) & "  type=" & o.GetButtonType(b)
    Next
  End If
  Err.Clear
  On Error GoTo 0
End Sub

' lists every toolbar/shell found by name inside each window (these are not in the normal tree)
Sub Snapshot(label)
  Dim nw, wi, w, shells, k, ids, i, o
  W ""
  W "==== " & label & " (" & Now & ") ===="
  nw = session.Children.Count
  W "windows open: " & nw
  For wi = 0 To nw - 1
    Set w = session.Children(CLng(wi))
    W "window " & wi & ": " & w.Id & "  text=" & w.Text
    On Error Resume Next
    Set shells = w.findAllByName("shell", "GuiShell")
    If Err.Number = 0 Then
      W "  shells found: " & shells.Count
      For k = 0 To shells.Count - 1
        DescribeShell shells(CLng(k))
      Next
    Else
      W "  findAllByName failed: " & Err.Description
    End If
    Err.Clear
    On Error GoTo 0
  Next
  ids = Array("wnd[0]/shellcont", "wnd[0]/shellcont/shell", "wnd[0]/shellcont[0]", "wnd[0]/shellcont[0]/shell", "wnd[0]/shellcont[1]", "wnd[0]/shellcont[1]/shell", "wnd[0]/titl/shellcont", "wnd[0]/titl/shellcont/shell", "wnd[1]/shellcont", "wnd[1]/shellcont/shell", "wnd[1]/titl/shellcont/shell", "wnd[2]/titl/shellcont/shell")
  For i = 0 To UBound(ids)
    Set o = Find(ids(i))
    If Not o Is Nothing Then
      W "probe " & ids(i) & ": EXISTS"
      DescribeShell o
    End If
  Next
End Sub

On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
  MsgBox "SAP GUI is not running.", 16, "Explore Services for Object"
  WScript.Quit 1
End If
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children(0)
Set session = connection.Children(0)
If Err.Number <> 0 Then
  MsgBox "Could not attach to a logged-in SAP session.", 16, "Explore Services for Object"
  WScript.Quit 1
End If
On Error GoTo 0

W "screen " & session.info.screenNumber & "  program " & session.info.program & "  title: " & session.findById("wnd[0]").Text
Snapshot "BEFORE the click"

Dim m
Set m = Find("wnd[0]/mbar/menu[5]/menu[6]")
If m Is Nothing Then
  W "menu System > Services for Object NOT FOUND"
Else
  W ""
  W "clicking System > Services for Object once: text=" & m.Text
  On Error Resume Next
  m.select
  If Err.Number <> 0 Then
    W "select failed: " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0
  WScript.Sleep 5000
  Snapshot "5 s after the click"
  WScript.Sleep 15000
  Snapshot "20 s after the click"
  WScript.Sleep 20000
  Snapshot "40 s after the click"
End If
outFile.Close
MsgBox "Done. Results are in gos-explore.txt (same folder as this script)." & vbCrLf & "Close any window SAP opened, and do NOT press anything else in SAP while the script runs.", 64, "Explore Services for Object"
