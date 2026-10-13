' SAP diagnose: shows what SAP exposes to scripts on this PC. Changes nothing in SAP.
' Log on to SAP first (Easy Access screen), then double-click this file and send the result file / popup text.
Option Explicit
Dim fso, sh, out, t, sapgui, app, i, j, conn, sess, wmi, procs, p, own, owner, dom, usr, r
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
t = ""
Sub L(s)
  t = t & s & vbCrLf
End Sub

On Error Resume Next
L "Time: " & Now
L "Windows user: " & sh.ExpandEnvironmentStrings("%USERDOMAIN%\%USERNAME%")
L "Script engine: " & WScript.FullName & " (" & ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & ")"
L "64-bit Windows: " & (sh.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%") & " / " & sh.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%"))

' SAP processes and who owns them
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
Set procs = wmi.ExecQuery("select * from Win32_Process where Name='saplogon.exe' or Name='sapgui.exe' or Name='saplgpad.exe' or Name='sapfewp.exe' or Name='NWBC.exe' or Name='wscript.exe'")
For Each p In procs
  owner = "": dom = ""
  r = p.GetOwner(owner, dom)
  L "Process " & p.Name & " pid " & p.ProcessId & " session " & p.SessionId & " owner " & dom & "\" & owner
Next
Err.Clear

Set sapgui = Nothing
Set sapgui = GetObject("SAPGUI")
If Err.Number <> 0 Then
  L "GetObject(SAPGUI) FAILED: " & Err.Number & " " & Err.Description
  Err.Clear
Else
  L "GetObject(SAPGUI): ok"
  Set app = sapgui.GetScriptingEngine
  If Err.Number <> 0 Then
    L "GetScriptingEngine FAILED: " & Err.Number & " " & Err.Description
    Err.Clear
  Else
    L "SAP GUI version: " & app.MajorVersion & "." & app.MinorVersion & " revision " & app.Revision & " patch " & app.PatchLevel
    Err.Clear
    L "Connections (Children): " & app.Children.Count
    For i = 0 To app.Children.Count - 1
      Set conn = app.Children(CLng(i))
      L "  Connection " & i & ": " & conn.Description & " | id " & conn.Id & " | sessions: " & conn.Children.Count
      If Err.Number <> 0 Then L "    error: " & Err.Number & " " & Err.Description
      Err.Clear
      For j = 0 To conn.Children.Count - 1
        Set sess = conn.Children(CLng(j))
        L "    Session " & j & ": user [" & sess.Info.User & "] system " & sess.Info.SystemName & " client " & sess.Info.Client & " language " & sess.Info.Language & " transaction " & sess.Info.Transaction
        If Err.Number <> 0 Then L "    error: " & Err.Number & " " & Err.Description
        Err.Clear
      Next
    Next
    Set sess = Nothing
    Set sess = app.ActiveSession
    If Err.Number <> 0 Then
      L "ActiveSession error: " & Err.Number & " " & Err.Description
      Err.Clear
    ElseIf sess Is Nothing Then
      L "ActiveSession: Nothing"
    Else
      L "ActiveSession: " & sess.Id & " user " & sess.Info.User
    End If
    Err.Clear
    L "ConnectionErrorText: " & app.ConnectionErrorText
    Err.Clear
  End If
End If

Set out = fso.CreateTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "sap-diagnose-result.txt"), True)
out.Write t
out.Close
MsgBox t & vbCrLf & "(also saved as sap-diagnose-result.txt next to this file)", 64, "SAP diagnose"
