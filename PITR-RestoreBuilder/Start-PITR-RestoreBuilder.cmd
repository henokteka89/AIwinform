@echo off
rem Starts PITR-RestoreBuilder.ps1 (same folder) with no console window behind it.
rem Read-only by default. For the optional TDE certificate transfer (temporary SMB share), use Start-PITR-RestoreBuilder-Admin.cmd instead.
set "PS1=%~dp0PITR-RestoreBuilder.ps1"
if not exist "%PS1%" (
  echo PITR-RestoreBuilder.ps1 not found in %~dp0
  pause
  exit /b 1
)
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%"
