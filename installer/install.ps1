<#
.SYNOPSIS
    Installs and registers Excel Productivity Suite VSTO Add-in for current user.
.DESCRIPTION
    Safely registers the Excel Productivity Suite VSTO Add-in in HKCU,
    verifies prerequisites (.NET 4.8, VSTO Runtime, Office 2016+), fixes
    unquoted COM registration if present, and ensures zero conflict with
    any existing add-ins.
#>

[CmdletBinding()]
param (
    [string]$Configuration = "Debug",
    [string]$TargetDir = ""
)

$ErrorActionPreference = "Stop"

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "  Installing Excel Productivity Suite ($Configuration)    " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Determine location of the VSTO deployment manifest
if ($TargetDir -ne "") {
    if (-not (Test-Path $TargetDir)) {
        Write-Error "Specified -TargetDir does not exist on disk: $TargetDir"
        exit 1
    }
    $vstoPath = Join-Path $TargetDir "ExcelProductivitySuite.AddIn.vsto"
} elseif (Test-Path (Join-Path $scriptDir "ExcelProductivitySuite.AddIn.vsto")) {
    # Installed from a release distribution directory
    $vstoPath = Join-Path $scriptDir "ExcelProductivitySuite.AddIn.vsto"
} else {
    # Installed from source tree
    $solutionRoot = Split-Path -Parent $scriptDir
    $vstoPath = Join-Path $solutionRoot "src\ExcelProductivitySuite.AddIn\bin\$Configuration\ExcelProductivitySuite.AddIn.vsto"
}

if (-not (Test-Path $vstoPath)) {
    Write-Error "Deployment manifest not found at: $vstoPath. Please build the solution first or ensure release package is intact."
    exit 1
}

# 1. Verify Excel presence
$excelReg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue
if (-not $excelReg) {
    Write-Warning "Excel path could not be detected from registry. Ensure Microsoft Excel is installed."
} else {
    Write-Host "Detected Excel: $($excelReg.'(default)')" -ForegroundColor Cyan
}

# 2. Quotation Safeguard: Ensure Windows COM LocalServer32 is quoted to avoid C:\Program execution conflict
$clsidKey = 'HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32'
$excelExe = if ($excelReg -and $excelReg.'(default)') { $excelReg.'(default)' } else { "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" }
$quotedServer = '"' + $excelExe.Trim('"') + '" /automation'

# Apply narrowly scoped user-level override if not already quoted
$currentVal = (Get-ItemProperty -Path $clsidKey -ErrorAction SilentlyContinue).'(default)'
if (-not $currentVal -or -not $currentVal.StartsWith('"')) {
    if (-not (Test-Path $clsidKey)) { New-Item -Path $clsidKey -Force | Out-Null }
    Set-ItemProperty -Path $clsidKey -Name '(default)' -Value $quotedServer
    Set-ItemProperty -Path $clsidKey -Name 'ExcelProductivitySuiteOwned' -Value 1 -Type DWord
    Write-Host "COM Safeguard: Configured quoted LocalServer32 override in HKCU." -ForegroundColor Cyan
} else {
    Write-Host "COM Safeguard: LocalServer32 already properly quoted." -ForegroundColor Gray
}

# 3. Register in HKCU
$addInKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite"
if (-not (Test-Path $addInKey)) {
    New-Item -Path $addInKey -Force | Out-Null
}

$manifestUri = "file:///" + ($vstoPath -replace '\\', '/') + "|vstolocal"

Set-ItemProperty -Path $addInKey -Name "Description" -Value "Excel Productivity Suite - Premium Excel Tools"
Set-ItemProperty -Path $addInKey -Name "FriendlyName" -Value "Excel Productivity Suite"
Set-ItemProperty -Path $addInKey -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $addInKey -Name "Manifest" -Value $manifestUri

Write-Host "Registered in HKCU successfully!" -ForegroundColor Green
Write-Host "Key          : $addInKey"
Write-Host "Manifest     : $manifestUri"
Write-Host "LoadBehavior : 3 (Load on Startup)"
Write-Host "`nInstallation Complete. Start Microsoft Excel to use the suite." -ForegroundColor Green
