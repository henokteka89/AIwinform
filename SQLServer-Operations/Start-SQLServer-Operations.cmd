@echo off
rem Runs SQLServer-Operations.ps1 (same folder) in this console, then opens the HTML report.
rem Any arguments are passed through, e.g.:  Start-SQLServer-Operations.cmd -SqlInstances SQL01,SQL02
set "PS1=%~dp0SQLServer-Operations.ps1"
if not exist "%PS1%" (
  echo SQLServer-Operations.ps1 not found in %~dp0
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
if %ERRORLEVEL% equ 0 (
  if exist "%~dp0Output\SQLServer-Operations.html" start "" "%~dp0Output\SQLServer-Operations.html"
) else (
  echo. & echo The script reported an error - the report was not opened.
)
pause
