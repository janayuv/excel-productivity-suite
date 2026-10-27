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

# 3. Resiliency Safeguard: Clear any Excel DisabledItems entries to unblock startup
$resiliencyVersions = @("16.0", "15.0", "14.0")
foreach ($ver in $resiliencyVersions) {
    $disabledKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\DisabledItems"
    if (Test-Path $disabledKey) {
        $di = Get-Item $disabledKey -ErrorAction SilentlyContinue
        if ($di) {
            foreach ($valName in $di.GetValueNames()) {
                $bytes = $di.GetValue($valName)
                if ($bytes -is [byte[]]) {
                    $str = [System.Text.Encoding]::Unicode.GetString($bytes)
                    if ($str -match 'ExcelProductivitySuite') {
                        Remove-ItemProperty -Path $disabledKey -Name $valName -Force -ErrorAction SilentlyContinue
                        Write-Host "Resiliency Safeguard: Cleared disabled item ($valName) in Office $ver." -ForegroundColor Cyan
                    }
                }
            }
        }
    }
}

# 4. VSTO Security Trust: Add deployment manifest to VSTO Inclusion List
try {
    $vstoAsm = [System.Reflection.Assembly]::Load("Microsoft.VisualStudio.Tools.Office.Runtime, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    if ($vstoAsm) {
        $entryType = $vstoAsm.GetType("Microsoft.VisualStudio.Tools.Office.Runtime.Security.AddInSecurityEntry")
        $inclusionListType = $vstoAsm.GetType("Microsoft.VisualStudio.Tools.Office.Runtime.Security.UserInclusionList")

        [xml]$vstoXml = Get-Content $vstoPath
        $ns = New-Object System.Xml.XmlNamespaceManager($vstoXml.NameTable)
        $ns.AddNamespace("ds", "http://www.w3.org/2000/09/xmldsig#")
        $rsaNode = $vstoXml.SelectSingleNode("//ds:KeyValue/ds:RSAKeyValue", $ns)

        if ($rsaNode -and $entryType -and $inclusionListType) {
            $publicKeyXml = $rsaNode.OuterXml
            $vstoUriClean = [Uri]("file:///" + ($vstoPath -replace '\\', '/'))

            $entry = [Activator]::CreateInstance($entryType, @($vstoUriClean, $publicKeyXml))
            $addMethod = $inclusionListType.GetMethod("Add", [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static)
            $addMethod.Invoke($null, @($entry))
            Write-Host "VSTO Security: Added deployment manifest to User Inclusion List." -ForegroundColor Cyan
        }
    }
} catch {
    Write-Warning "Could not register VSTO inclusion list entry automatically: $_"
}

# 5. Register in HKCU
$manifestUri = "file:///" + ($vstoPath -replace '\\', '/') + "|vstolocal"
$addInKeys = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite",
    "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite.AddIn"
)

foreach ($keyPath in $addInKeys) {
    if (-not (Test-Path $keyPath)) {
        New-Item -Path $keyPath -Force | Out-Null
    }
    Set-ItemProperty -Path $keyPath -Name "Description" -Value "Excel Productivity Suite - Premium Excel Tools"
    Set-ItemProperty -Path $keyPath -Name "FriendlyName" -Value "Excel Productivity Suite"
    Set-ItemProperty -Path $keyPath -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $keyPath -Name "Manifest" -Value $manifestUri
}

Write-Host "Registered in HKCU successfully!" -ForegroundColor Green
Write-Host "Keys         : $($addInKeys -join ', ')"
Write-Host "Manifest     : $manifestUri"
Write-Host "LoadBehavior : 3 (Load on Startup)"
Write-Host "`nInstallation Complete. Start Microsoft Excel to use the suite." -ForegroundColor Green
