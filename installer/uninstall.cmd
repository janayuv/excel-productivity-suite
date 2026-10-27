@echo off
setlocal
cd /d "%~dp0"
echo ==========================================================
echo   Uninstalling Excel Productivity Suite
echo ==========================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
if %ERRORLEVEL% neq 0 (
    echo.
    echo Uninstallation encountered an error (Code: %ERRORLEVEL%).
)
pause
