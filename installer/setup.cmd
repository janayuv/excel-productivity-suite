@echo off
setlocal
cd /d "%~dp0"
echo ==========================================================
echo   Installing Excel Productivity Suite to Local Programs
echo ==========================================================
echo.
echo [1/2] Unblocking package files...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -Path '%~dp0' -Recurse | Unblock-File -ErrorAction SilentlyContinue"

echo [2/2] Deploying to %LOCALAPPDATA%\Programs\ExcelProductivitySuite and registering...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -InstallToLocalPrograms %*
if %ERRORLEVEL% neq 0 (
    echo.
    echo ==========================================================
    echo Installation encountered an error (Code: %ERRORLEVEL%).
    echo Run diagnostics.cmd to check prerequisites.
    echo ==========================================================
)
echo.
pause
