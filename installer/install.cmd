@echo off
setlocal
cd /d "%~dp0"
echo ==========================================================
echo   Installing Excel Productivity Suite
echo ==========================================================
echo.
echo [1/2] Unblocking package files...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File -ErrorAction SilentlyContinue"

echo [2/2] Registering add-in and applying Excel settings...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
if %ERRORLEVEL% neq 0 (
    echo.
    echo ==========================================================
    echo Installation encountered an error (Code: %ERRORLEVEL%).
    echo Run diagnostics.cmd to check prerequisites.
    echo ==========================================================
)
echo.
pause
