@echo off
rem Creates a Windows scheduled task that runs export-vendor-requests.vbs on weekdays at 08:05.
rem Keep this file in the same folder as the .vbs, then double-click it.
set "SCRIPT=%~dp0export-vendor-requests.vbs"
if not exist "%SCRIPT%" (
  echo Could not find "%SCRIPT%"
  pause
  exit /b 1
)
schtasks /create /tn "AVSF SAP export" /tr "wscript.exe \"%SCRIPT%\"" /sc weekly /d MON,TUE,WED,THU,FRI /st 08:05 /rl LIMITED /f
if errorlevel 1 (
  echo.
  echo The task could not be created. Right-click this file and choose "Run as administrator", or ask IT.
) else (
  echo.
  echo Done. "AVSF SAP export" will run Monday to Friday at 08:05 while you are logged in.
  echo SAP must be open and logged in, and the PC unlocked.
  echo To remove it later: schtasks /delete /tn "AVSF SAP export" /f
)
pause
