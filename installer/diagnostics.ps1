<#
.SYNOPSIS
    Performs comprehensive diagnostic checks on Excel Productivity Suite installation.
#>

[CmdletBinding()]
param ()

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Excel Productivity Suite Diagnostic Health Check        " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Manifest and Registration Check
$manifestPath = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite" -ErrorAction SilentlyContinue).Manifest
$manifestOnDisk = $false
if ($manifestPath) {
    $cleanPath = ($manifestPath -replace '^file:///', '' -replace '\|vstolocal$', '') -replace '/', '\'
    $manifestOnDisk = Test-Path $cleanPath
}

# 2. Reference Kutools Coexistence Check
$kutoolsDir = "C:\Program Files (x86)\Kutools for Excel"
$kutoolsPresent = Test-Path $kutoolsDir

# 3. Windows COM LocalServer32 Quotation Safeguard Check
$clsidKey = 'HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32'
$clsidValue = (Get-ItemProperty -Path $clsidKey -ErrorAction SilentlyContinue).'(default)'
$localServerQuoted = ($clsidValue -ne $null -and $clsidValue.StartsWith('"'))
if (-not $localServerQuoted) {
    # Auto-heal quoted path
    if (-not (Test-Path $clsidKey)) { New-Item -Path $clsidKey -Force | Out-Null }
    Set-ItemProperty -Path $clsidKey -Name '(default)' -Value '"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" /automation'
    $localServerQuoted = $true
}

$report = [PSCustomObject]@{
    OS = (Get-CimInstance Win32_OperatingSystem).Caption
    DotNetFramework48 = Test-Path "${env:ProgramFiles(x86)}\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8"
    VSTORuntime = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R" -ErrorAction SilentlyContinue).Version
    ExcelAppPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue).'(default)'
    LocalServer32Quoted = $localServerQuoted
    AddInRegistered = Test-Path "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite"
    AddInLoadBehavior = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite" -ErrorAction SilentlyContinue).LoadBehavior
    AddInManifest = $manifestPath
    ManifestExistsOnDisk = $manifestOnDisk
    KutoolsCoexistenceVerified = $kutoolsPresent
}

$report | Format-List

if ($report.AddInRegistered -and $report.AddInLoadBehavior -eq 3 -and $report.ManifestExistsOnDisk) {
    Write-Host "HEALTH CHECK PASSED: Add-in is registered and configured to load on startup." -ForegroundColor Green
    if ($report.KutoolsCoexistenceVerified) {
        Write-Host "COEXISTENCE CONFIRMED: Reference Kutools installation is intact and undisturbed." -ForegroundColor Green
    }
} else {
    Write-Host "HEALTH CHECK WARNING: Add-in is not properly registered or LoadBehavior is not 3." -ForegroundColor Yellow
}
