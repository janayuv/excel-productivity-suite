@echo off
setlocal
cd /d "%~dp0"
echo ==========================================================
echo   Installing Excel Productivity Suite
echo ==========================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
if %ERRORLEVEL% neq 0 (
    echo.
    echo Installation encountered an error (Code: %ERRORLEVEL%).
)
pause
