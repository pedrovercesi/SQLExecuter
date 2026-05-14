@echo off
REM Double-click wrapper for Apply-Reports.ps1
REM Runs the PowerShell script with execution policy bypass and keeps the window open at the end.
setlocal
set "SCRIPT_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Apply-Reports.ps1"
echo.
echo --- Done. Press any key to close ---
pause >nul
endlocal
