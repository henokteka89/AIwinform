@echo off
rem Starts SqlFleetStudio.ps1 (same folder) with no console window behind it.
set "PS1=%~dp0SqlFleetStudio.ps1"
if not exist "%PS1%" (
  echo SqlFleetStudio.ps1 not found in %~dp0
  pause
  exit /b 1
)
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%" %*
