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
    [string]$Version = "5.8.0",
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
if (-not (Test-Path $stageDir)) {
    New-Item -Path $stageDir -ItemType Directory -Force | Out-Null
} else {
    Get-ChildItem -Path $stageDir | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

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

# Copy installer scripts and tools
$installerFiles = @(
    "install.ps1",
    "uninstall.ps1",
    "diagnostics.ps1",
    "install.cmd",
    "setup.cmd",
    "uninstall.cmd",
    "diagnostics.cmd",
    "ExcelProductivitySuite.cer"
)
foreach ($iFile in $installerFiles) {
    $srcFile = Join-Path $scriptDir $iFile
    if (Test-Path $srcFile) {
        Copy-Item -Path $srcFile -Destination $stageDir -Force
        Write-Host "  -> Copied installer component: $iFile" -ForegroundColor Gray
    }
}

# Create release notice and user guide for other PCs
$readmeContent = @"
========================================================================
Excel Productivity Suite - Version $Version (Windows Desktop x64 / x86)
Host Application : Microsoft Excel 2016 / 2019 / 2021 / Microsoft 365
Target Runtime   : .NET Framework 4.8 & Visual Studio Tools for Office
========================================================================

PREREQUISITES FOR TARGET PC:
1. Windows 10 (version 1903+) or Windows 11 (or Win 7/8.1 with .NET 4.8).
2. Microsoft Excel 2016, 2019, 2021, or Microsoft 365 (32-bit or 64-bit).
3. Visual Studio Tools for Office (VSTO) Runtime:
   Included by default with standard Microsoft Office installations.
   (If missing, download free from: https://www.microsoft.com/en-us/download/details.aspx?id=48217)

HOW TO INSTALL ON THIS PC:

Method 1: Permanent Setup (Recommended)
   - Extract the entire ZIP archive to any temporary folder (e.g. Downloads).
   - Double-click 'setup.cmd'.
   - This automatically unblocks files, registers the code-signing certificate,
     installs the suite to '%LOCALAPPDATA%\Programs\ExcelProductivitySuite',
     and configures Excel.
   - After installation, you can safely delete the extracted folder.

Method 2: In-Place / Portable Installation
   - Place this folder in your desired permanent location (e.g. C:\Tools\ExcelProductivitySuite).
   - Double-click 'install.cmd'.
   - The add-in will be registered directly from this folder.

VERIFY INSTALLATION:
   - Double-click 'diagnostics.cmd'.
   - All items should show PASSED.
   - Open Microsoft Excel. You will see two new ribbon tabs:
     [Excel Productivity] and [Workbook & Sheets].
   - Press Ctrl+Shift+P anytime inside Excel to open the Quick Tool Launcher (SuperSearch).

UNINSTALL:
   - Double-click 'uninstall.cmd' to cleanly unregister the add-in from Excel.

TROUBLESHOOTING:
   - If Excel was already open during installation, close and reopen Excel.
   - If Excel shows a disabled items prompt, run 'diagnostics.cmd' to automatically detect
     and clear any disabled flags.
========================================================================
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
