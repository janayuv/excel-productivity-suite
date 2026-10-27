<#
.SYNOPSIS
    Performs comprehensive diagnostic checks on Excel Productivity Suite installation.
#>

[CmdletBinding()]
param ()

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Excel Productivity Suite Diagnostic Health Check        " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

$report = [PSCustomObject]@{
    OS = (Get-CimInstance Win32_OperatingSystem).Caption
    DotNetFramework48 = Test-Path "${env:ProgramFiles(x86)}\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8"
    VSTORuntime = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R" -ErrorAction SilentlyContinue).Version
    ExcelAppPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue).'(default)'
    AddInRegistered = Test-Path "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite"
    AddInLoadBehavior = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite" -ErrorAction SilentlyContinue).LoadBehavior
    AddInManifest = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite" -ErrorAction SilentlyContinue).Manifest
}

$report | Format-List

if ($report.AddInRegistered -and $report.AddInLoadBehavior -eq 3) {
    Write-Host "HEALTH CHECK PASSED: Add-in is registered and configured to load on startup." -ForegroundColor Green
} else {
    Write-Host "HEALTH CHECK WARNING: Add-in is not properly registered or LoadBehavior is not 3." -ForegroundColor Yellow
}
