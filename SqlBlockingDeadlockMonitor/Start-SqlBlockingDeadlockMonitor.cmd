@echo off
rem Starts SqlBlockingDeadlockMonitor.ps1 (same folder) with no console window behind it.
set "PS1=%~dp0SqlBlockingDeadlockMonitor.ps1"
if not exist "%PS1%" (
  echo SqlBlockingDeadlockMonitor.ps1 not found in %~dp0
  pause
  exit /b 1
)
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%"
