@echo off
rem Starts SqlWaitMonitor.ps1 (same folder) for a CENTRAL install (database DBAWaitCentral, many servers via linked servers),
rem with no console window behind it. Set SERVER to the CENTRAL server; pick the monitored server inside the window.
set "SERVER=localhost"
set "PS1=%~dp0SqlWaitMonitor.ps1"
if not exist "%PS1%" (
  echo SqlWaitMonitor.ps1 not found in %~dp0
  pause
  exit /b 1
)
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%" -Server "%SERVER%" -Database DBAWaitCentral %*
