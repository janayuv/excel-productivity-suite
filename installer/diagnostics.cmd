@echo off
setlocal
cd /d "%~dp0"
echo ==========================================================
echo   Running Excel Productivity Suite Diagnostics
echo ==========================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0diagnostics.ps1" %*
echo.
pause
