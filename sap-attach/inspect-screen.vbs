' READ-ONLY inspector: writes the objects of the current SAP screen (and any popups / floating windows)
' to sap-screen-tree.txt next to this script. It does not press, type or change anything in SAP.
Option Explicit

Dim fso, outFile, session, application, connection, SapGuiAuto, lineCount, MAX_LINES, MAX_DEPTH
MAX_LINES = 2500
MAX_DEPTH = 9
lineCount = 0
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "sap-screen-tree.txt"), True, True)

Sub W(txt)
  If lineCount < MAX_LINES Then outFile.WriteLine txt
  lineCount = lineCount + 1
End Sub

Sub Walk(obj, depth)
  Dim i, n, c, line, t, bc, b
  If depth > MAX_DEPTH Or lineCount > MAX_LINES Then Exit Sub
  On Error Resume Next
  line = Space(depth * 2) & obj.Id & "  |  " & obj.Type
  t = ""
  t = obj.SubType
  If Err.Number = 0 And t <> "" Then line = line & " / " & t
  Err.Clear
  t = ""
  t = obj.Text
  If Err.Number = 0 And t <> "" Then line = line & "  |  text=" & Left(t, 60)
  Err.Clear
  t = ""
  t = obj.Tooltip
  If Err.Number = 0 And t <> "" Then line = line & "  |  tip=" & Left(t, 60)
  Err.Clear
  W line
  ' toolbars: list their buttons
  bc = -1
  bc = obj.ButtonCount
  If Err.Number = 0 And bc >= 0 Then
    For b = 0 To bc - 1
      W Space(depth * 2 + 4) & "button " & b & ": id=" & obj.GetButtonId(b) & "  text=" & obj.GetButtonText(b) & "  tip=" & obj.GetButtonTooltip(b) & "  type=" & obj.GetButtonType(b)
    Next
  End If
  Err.Clear
  n = obj.Children.Count
  If Err.Number = 0 Then
    For i = 0 To n - 1
      Set c = Nothing
      Set c = obj.Children(CLng(i))
      If Not c Is Nothing Then Walk c, depth + 1
    Next
  End If
  Err.Clear
  On Error GoTo 0
End Sub

On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
  MsgBox "SAP GUI is not running.", 16, "Inspect SAP screen"
  WScript.Quit 1
End If
Set application = SapGuiAuto.GetScriptingEngine
Set connection = application.Children(0)
Set session = connection.Children(0)
If Err.Number <> 0 Then
  MsgBox "Could not attach to a logged-in SAP session.", 16, "Inspect SAP screen"
  WScript.Quit 1
End If
On Error GoTo 0

W "Screen: " & session.info.screenNumber & "  program: " & session.info.program & "  transaction: " & session.info.transaction
Dim wi, nw
nw = session.Children.Count
For wi = 0 To nw - 1
  W "---- window " & wi & " ----"
  Walk session.Children(CLng(wi)), 0
Next
outFile.Close
MsgBox "Done. " & lineCount & " lines written to sap-screen-tree.txt (same folder as this script).", 64, "Inspect SAP screen"
