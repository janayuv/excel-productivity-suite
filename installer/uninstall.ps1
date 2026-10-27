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

$addInKeys = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite",
    "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite.AddIn"
)

$manifestPathToClean = $null
foreach ($keyPath in $addInKeys) {
    if (Test-Path $keyPath) {
        $reg = Get-ItemProperty -Path $keyPath -ErrorAction SilentlyContinue
        if ($reg -and $reg.Manifest -and -not $manifestPathToClean) {
            $manifestPathToClean = ($reg.Manifest -replace '\|vstolocal$', '')
        }
        Remove-Item -Path $keyPath -Recurse -Force
        Write-Host "Successfully removed registration from: $keyPath" -ForegroundColor Green
    }
}

# Clean VSTO Inclusion List entry if uri is known
if ($manifestPathToClean) {
    try {
        $vstoAsm = [System.Reflection.Assembly]::Load("Microsoft.VisualStudio.Tools.Office.Runtime, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
        if ($vstoAsm) {
            $inclusionListType = $vstoAsm.GetType("Microsoft.VisualStudio.Tools.Office.Runtime.Security.UserInclusionList")
            $removeMethod = $inclusionListType.GetMethod("Remove", [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static, $null, @([Uri]), $null)
            $removeMethod.Invoke($null, @([Uri]$manifestPathToClean))
            Write-Host "VSTO Security: Cleaned User Inclusion List entry." -ForegroundColor Green
        }
    } catch {
        # Inclusion cleanup is best-effort
    }
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
