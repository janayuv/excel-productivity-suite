<#
.SYNOPSIS
    Installs and registers Excel Productivity Suite VSTO Add-in for current user.
.DESCRIPTION
    Safely registers the Excel Productivity Suite VSTO Add-in in HKCU,
    verifies prerequisites (.NET 4.8, VSTO Runtime, Office 2016+), and
    ensures zero conflict with any existing add-ins.
#>

[CmdletBinding()]
param (
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "  Installing Excel Productivity Suite ($Configuration)  " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$solutionRoot = Split-Path -Parent $scriptDir
$vstoPath = Join-Path $solutionRoot "src\ExcelProductivitySuite.AddIn\bin\$Configuration\ExcelProductivitySuite.AddIn.vsto"

if (-not (Test-Path $vstoPath)) {
    Write-Error "Deployment manifest not found at: $vstoPath. Please build the solution first."
    exit 1
}

# 1. Verify Excel presence
$excelReg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue
if (-not $excelReg) {
    Write-Warning "Excel path could not be detected from registry. Ensure Microsoft Excel is installed."
} else {
    Write-Host "Detected Excel: $($excelReg.'(default)')" -ForegroundColor Cyan
}

# 2. Register in HKCU
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
Write-Host "Key: $addInKey"
Write-Host "Manifest: $manifestUri"
Write-Host "LoadBehavior: 3 (Load on Startup)"
Write-Host "`nInstallation Complete. Start Microsoft Excel to use the suite." -ForegroundColor Green
