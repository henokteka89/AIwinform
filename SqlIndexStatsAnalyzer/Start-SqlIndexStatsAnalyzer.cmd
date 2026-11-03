@echo off
rem Starts SqlIndexStatsAnalyzer.ps1 (same folder) with no console window behind it.
rem Optional: set SERVER (and DATABASE) to connect straight away, e.g. PRODSQL01 or HOST\INSTANCE or HOST,PORT.
set "SERVER="
set "DATABASE="
set "PS1=%~dp0SqlIndexStatsAnalyzer.ps1"
if not exist "%PS1%" (
  echo SqlIndexStatsAnalyzer.ps1 not found in %~dp0
  pause
  exit /b 1
)
if defined SERVER (
  if defined DATABASE (
    start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%" -Server "%SERVER%" -Database "%DATABASE%" %*
  ) else (
    start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%" -Server "%SERVER%" %*
  )
) else (
  start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%" %*
)
