@echo off
rem Starts SqlPatchGui.ps1 (same folder) elevated, with no console window behind it.
set "PS1=%~dp0SqlPatchGui.ps1"
if not exist "%PS1%" (
  echo SqlPatchGui.ps1 not found in %~dp0
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-WindowStyle','Hidden','-File',('\"' + $env:PS1 + '\"'))"
