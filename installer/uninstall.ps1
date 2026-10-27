<#
.SYNOPSIS
    Unregisters Excel Productivity Suite VSTO Add-in cleanly from current user.
#>

[CmdletBinding()]
param ()

$ErrorActionPreference = "Stop"

Write-Host "==========================================================" -ForegroundColor Yellow
Write-Host "  Uninstalling Excel Productivity Suite                 " -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Yellow

$addInKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite"
if (Test-Path $addInKey) {
    Remove-Item -Path $addInKey -Recurse -Force
    Write-Host "Successfully removed registration from: $addInKey" -ForegroundColor Green
} else {
    Write-Host "Registration key was not found. Nothing to remove." -ForegroundColor Gray
}

Write-Host "`nUninstallation Complete. Excel will no longer load this add-in." -ForegroundColor Green
