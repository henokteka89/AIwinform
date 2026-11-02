@echo off
rem Starts SqlWaitMonitor.ps1 (same folder) for a STANDALONE install (database DBAWait), with no console window behind it.
rem Set SERVER to the monitored SQL Server, e.g. PRODSQL01 or HOST\INSTANCE or HOST,PORT.
set "SERVER=localhost"
set "PS1=%~dp0SqlWaitMonitor.ps1"
if not exist "%PS1%" (
  echo SqlWaitMonitor.ps1 not found in %~dp0
  pause
  exit /b 1
)
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%" -Server "%SERVER%" -Database DBAWait %*
