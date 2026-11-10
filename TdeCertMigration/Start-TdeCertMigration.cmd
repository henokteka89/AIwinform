@echo off
rem Starts TdeCertMigration.ps1 (same folder) with no console window behind it.
rem Opens in What-If mode: nothing is changed until you choose Execute and type the destination server name.
set "PS1=%~dp0TdeCertMigration.ps1"
if not exist "%PS1%" (
  echo TdeCertMigration.ps1 not found in %~dp0
  pause
  exit /b 1
)
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%"
