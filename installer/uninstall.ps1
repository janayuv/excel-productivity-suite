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

# Clean up owned COM LocalServer32 override if created by this installer
$clsidKey = 'HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32'
if (Test-Path $clsidKey) {
    $owned = (Get-ItemProperty -Path $clsidKey -Name 'ExcelProductivitySuiteOwned' -ErrorAction SilentlyContinue).ExcelProductivitySuiteOwned
    if ($owned -eq 1) {
        Remove-Item -Path $clsidKey -Recurse -Force -ErrorAction SilentlyContinue
        $parentClsid = 'HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}'
        if ((Get-ChildItem -Path $parentClsid -ErrorAction SilentlyContinue).Count -eq 0) {
            Remove-Item -Path $parentClsid -Force -ErrorAction SilentlyContinue
        }
        Write-Host "Cleaned up owned COM LocalServer32 quotation override." -ForegroundColor Green
    }
}

Write-Host "`nUninstallation Complete. Excel will no longer load this add-in." -ForegroundColor Green
