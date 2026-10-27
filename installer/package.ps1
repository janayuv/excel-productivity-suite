<#
.SYNOPSIS
    Builds, stages, and packages Excel Productivity Suite into a versioned release bundle.
.DESCRIPTION
    Compiles the solution in Release mode, collects production artifacts into dist/,
    generates install/uninstall scripts, computes SHA256 checksums, and produces
    a distributable zip package.
#>

[CmdletBinding()]
param (
    [string]$Version = "4.0.0",
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Packaging Excel Productivity Suite v$Version ($Configuration)  " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$solutionRoot = Split-Path -Parent $scriptDir
$distDir = Join-Path $solutionRoot "dist"
$stageDir = Join-Path $distDir "ExcelProductivitySuite-v$Version-win-x64"

# 1. Build Solution in Release Configuration
Write-Host "`n[1/5] Building solution in $Configuration configuration..." -ForegroundColor Yellow
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe"
if (-not (Test-Path $msbuild)) {
    $msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
}

$solutionPath = Join-Path $solutionRoot "ExcelProductivitySuite.sln"
& $msbuild $solutionPath /t:Build /p:Configuration=$Configuration /v:m

if ($LASTEXITCODE -ne 0) {
    Write-Error "Release build failed with exit code $LASTEXITCODE."
    exit 1
}
Write-Host "Release build succeeded cleanly." -ForegroundColor Green

# 2. Prepare staging directory
Write-Host "`n[2/5] Creating staging layout at: $stageDir..." -ForegroundColor Yellow
if (Test-Path $stageDir) {
    Remove-Item -Path $stageDir -Recurse -Force
}
New-Item -Path $stageDir -ItemType Directory -Force | Out-Null

$addInBin = Join-Path $solutionRoot "src\ExcelProductivitySuite.AddIn\bin\$Configuration"

# 3. Copy binaries and deployment manifests
Write-Host "`n[3/5] Copying binaries and manifests..." -ForegroundColor Yellow
$artifactsToCopy = @(
    "ExcelProductivitySuite.AddIn.dll",
    "ExcelProductivitySuite.AddIn.vsto",
    "ExcelProductivitySuite.Core.dll",
    "ExcelProductivitySuite.Engine.dll",
    "ExcelProductivitySuite.Infrastructure.dll",
    "ExcelProductivitySuite.UI.dll",
    "Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
)

foreach ($file in $artifactsToCopy) {
    $src = Join-Path $addInBin $file
    if (Test-Path $src) {
        Copy-Item -Path $src -Destination $stageDir -Force
        Write-Host "  -> Copied: $file" -ForegroundColor Gray
    } else {
        Write-Warning "File not found: $src"
    }
}

# Copy manifest file if present
$manifestPattern = Join-Path $addInBin "*.manifest"
Get-ChildItem -Path $manifestPattern -ErrorAction SilentlyContinue | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination $stageDir -Force
    Write-Host "  -> Copied manifest: $($_.Name)" -ForegroundColor Gray
}

# Copy Ribbon XML folder
$ribbonDir = Join-Path $stageDir "Ribbon"
New-Item -Path $ribbonDir -ItemType Directory -Force | Out-Null
$srcRibbon = Join-Path $solutionRoot "src\ExcelProductivitySuite.AddIn\Ribbon\ExcelProductivityRibbon.xml"
if (Test-Path $srcRibbon) {
    Copy-Item -Path $srcRibbon -Destination $ribbonDir -Force
    Write-Host "  -> Copied: Ribbon\ExcelProductivityRibbon.xml" -ForegroundColor Gray
}

# Copy installer scripts
Copy-Item -Path (Join-Path $scriptDir "install.ps1") -Destination $stageDir -Force
Copy-Item -Path (Join-Path $scriptDir "uninstall.ps1") -Destination $stageDir -Force
Copy-Item -Path (Join-Path $scriptDir "diagnostics.ps1") -Destination $stageDir -Force

# Create release notice
$readmeContent = @"
========================================================================
Excel Productivity Suite - Version $Version (Windows Desktop x64 / x86)
Host Application : Microsoft Excel 2016 / 2019 / 2021 / Microsoft 365
Target Runtime   : .NET Framework 4.8 & Visual Studio Tools for Office
========================================================================

QUICK INSTALLATION:
1. Open PowerShell as normal user (no Administrator privileges required).
2. Run:
   powershell.exe -ExecutionPolicy Bypass -File .\install.ps1
3. Launch Microsoft Excel. The 'Excel Productivity' and 'Workbook & Sheets'
   tabs will appear automatically on the Office Ribbon.

VERIFY INSTALLATION:
Run:
   powershell.exe -ExecutionPolicy Bypass -File .\diagnostics.ps1

UNINSTALL:
Run:
   powershell.exe -ExecutionPolicy Bypass -File .\uninstall.ps1

CODE SIGNING NOTE:
This release package is signed with the internal developer certificate
(Thumbprint: 74F1C194A7C449209F1795774588A5D07C4375D6). For enterprise
production deployments, replace with your organization's Authenticode CA.
"@
Set-Content -Path (Join-Path $stageDir "README.txt") -Value $readmeContent

# 4. Generate ZIP archive
Write-Host "`n[4/5] Creating distribution zip archive..." -ForegroundColor Yellow
$zipPath = Join-Path $distDir "ExcelProductivitySuite-v$Version-win-x64.zip"
if (Test-Path $zipPath) {
    Remove-Item -Path $zipPath -Force
}
Compress-Archive -Path "$stageDir\*" -DestinationPath $zipPath -Force
Write-Host "  -> Created archive: $zipPath" -ForegroundColor Green

# 5. Generate Checksum Manifest (SHA256)
Write-Host "`n[5/5] Generating SHA256 checksum manifest..." -ForegroundColor Yellow
$checksumFile = Join-Path $distDir "SHA256SUMS.txt"
$checksumLines = [System.Collections.Generic.List[string]]::new()

$zipHash = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash
$checksumLines.Add("$zipHash  ExcelProductivitySuite-v$Version-win-x64.zip")

Get-ChildItem -Path $stageDir -Recurse -File | ForEach-Object {
    $h = (Get-FileHash -Path $_.FullName -Algorithm SHA256).Hash
    $rel = $_.FullName.Substring($stageDir.Length + 1).Replace('\', '/')
    $checksumLines.Add("$h  $rel")
}

Set-Content -Path $checksumFile -Value $checksumLines
Write-Host "  -> Written checksums to: $checksumFile" -ForegroundColor Green

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "  RELEASE PACKAGING COMPLETE: v$Version ($Configuration)   " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Staging Directory : $stageDir"
Write-Host "Release Archive   : $zipPath"
Write-Host "Checksum Manifest : $checksumFile"
