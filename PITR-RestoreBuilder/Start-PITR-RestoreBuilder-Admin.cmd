@echo off
rem Starts PITR-RestoreBuilder.ps1 (same folder) ELEVATED, with no console window behind it.
rem Only needed for the optional TDE certificate transfer, which creates (and then removes) a temporary SMB share on this machine.
set "PS1=%~dp0PITR-RestoreBuilder.ps1"
if not exist "%PS1%" (
  echo PITR-RestoreBuilder.ps1 not found in %~dp0
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-WindowStyle','Hidden','-File',('\"' + $env:PS1 + '\"'))"
