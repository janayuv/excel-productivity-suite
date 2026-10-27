<#
.SYNOPSIS
    Automated Release Package & Installation Lifecycle Validation Suite for Phase 5.
.DESCRIPTION
    Validates standalone extraction of ExcelProductivitySuite-v4.0.0-win-x64.zip,
    installation outside repository, Excel COM startup, Kutools coexistence,
    uninstallation, reinstallation, upgrade, and error recovery.
#>

[CmdletBinding()]
param (
    [string]$ZipPath = "D:\Kutools-New\dist\ExcelProductivitySuite-v4.0.0-win-x64.zip",
    [string]$DeployBaseDir = "C:\Users\Yogeswari\AppData\Local\Programs\ExcelProductivitySuite-ReleaseTest"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

function Record-Step {
    param (
        [string]$ScenarioId,
        [string]$Title,
        [string]$Expected,
        [string]$Actual,
        [bool]$Passed,
        [double]$ElapsedMs
    )
    $obj = [PSCustomObject]@{
        ScenarioId = $ScenarioId
        Title = $Title
        Expected = $Expected
        Actual = $Actual
        Passed = $Passed
        ElapsedMs = [Math]::Round($ElapsedMs, 2)
    }
    $results.Add($obj)
    $color = if ($Passed) { "Green" } else { "Red" }
    $tag = if ($Passed) { "[PASS]" } else { "[FAIL]" }
    Write-Host "$tag $ScenarioId - $Title ($([Math]::Round($ElapsedMs, 2)) ms)" -ForegroundColor $color
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Phase 5 Release Package & Installation Lifecycle Tests  " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# -----------------------------------------------------------------
# S01: Extract Release ZIP to External Directory (Outside Repository)
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    if (Test-Path $DeployBaseDir) {
        Remove-Item -Path $DeployBaseDir -Recurse -Force
    }
    New-Item -Path $DeployBaseDir -ItemType Directory -Force | Out-Null
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $DeployBaseDir)
    $sw.Stop()
    $extractedFiles = Get-ChildItem -Path $DeployBaseDir -Recurse -File
    $s01Passed = ($extractedFiles.Count -ge 12 -and (Test-Path "$DeployBaseDir\ExcelProductivitySuite.AddIn.vsto"))
    Record-Step -ScenarioId "S01" -Title "Extract Release ZIP to External Directory" `
        -Expected "All files extracted cleanly outside repository" `
        -Actual ("Extracted " + $extractedFiles.Count + " files to " + $DeployBaseDir) `
        -Passed $s01Passed -ElapsedMs $sw.Elapsed.TotalMilliseconds
} catch {
    Record-Step -ScenarioId "S01" -Title "Extract Release ZIP to External Directory" -Expected "Success" -Actual $_.ToString() -Passed $false -ElapsedMs 0
}

# -----------------------------------------------------------------
# S02: Fresh Installation from External Production Location
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $installScript = Join-Path $DeployBaseDir "install.ps1"
    & powershell.exe -ExecutionPolicy Bypass -File $installScript
    $sw.Stop()

    $addInKey = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelProductivitySuite"
    $reg = Get-ItemProperty -Path $addInKey -ErrorAction SilentlyContinue
    $manifestMatches = $reg -ne $null -and $reg.Manifest.Contains($DeployBaseDir.Replace('\', '/'))
    $loadBehaviorOk = $reg -ne $null -and $reg.LoadBehavior -eq 3
    $noDevPath = $reg -ne $null -and (-not $reg.Manifest.Contains("Kutools-New"))

    $s02Passed = ($manifestMatches -and $loadBehaviorOk -and $noDevPath)
    Record-Step -ScenarioId "S02" -Title "Fresh Installation from External Directory" `
        -Expected "Manifest points to external location without dev path, LoadBehavior=3" `
        -Actual ("Manifest: " + $reg.Manifest + " | LoadBehavior: " + $reg.LoadBehavior) `
        -Passed $s02Passed -ElapsedMs $sw.Elapsed.TotalMilliseconds
} catch {
    Record-Step -ScenarioId "S02" -Title "Fresh Installation from External Directory" -Expected "Success" -Actual $_.ToString() -Passed $false -ElapsedMs 0
}

# -----------------------------------------------------------------
# S03: External Diagnostics Verification
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $diagScript = Join-Path $DeployBaseDir "diagnostics.ps1"
    $diagOutput = & powershell.exe -ExecutionPolicy Bypass -File $diagScript
    $sw.Stop()
    $diagPassed = ($diagOutput -match "HEALTH CHECK PASSED" -and $diagOutput -match "COEXISTENCE CONFIRMED")
    Record-Step -ScenarioId "S03" -Title "External Diagnostic Health Check" `
        -Expected "HEALTH CHECK PASSED and COEXISTENCE CONFIRMED" `
        -Actual $(if ($diagPassed) { "Health check & coexistence confirmed from external install" } else { "Diagnostic failed: $diagOutput" }) `
        -Passed $diagPassed -ElapsedMs $sw.Elapsed.TotalMilliseconds
} catch {
    Record-Step -ScenarioId "S03" -Title "External Diagnostic Health Check" -Expected "Success" -Actual $_.ToString() -Passed $false -ElapsedMs 0
}

# -----------------------------------------------------------------
# S04: First Excel Startup & COM Add-In Verification from External Location
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $beforePids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $afterPids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPid = $afterPids | Where-Object { $beforePids -notcontains $_ } | Select-Object -First 1

    $wb = $excel.Workbooks.Add()
    
    # Check COM Addins in running Excel
    $foundAddin = $null
    for ($i = 1; $i -le $excel.COMAddIns.Count; $i++) {
        $comItem = $excel.COMAddIns.Item($i)
        if ($comItem.ProgId -eq "ExcelProductivitySuite") {
            $foundAddin = $comItem
            break
        }
    }
    
    $isRegistered = ($foundAddin -ne $null)
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if ($newPid) {
        $proc = Get-Process -Id $newPid -ErrorAction SilentlyContinue
        if ($proc) {
            if (-not $proc.WaitForExit(3000)) { $proc.Kill(); $proc.WaitForExit(2000) }
        }
    }
    $sw.Stop()

    $s04Passed = $isRegistered
    Record-Step -ScenarioId "S04" -Title "First Excel Startup & COM AddIn Registration" `
        -Expected "Excel recognizes ExcelProductivitySuite COM add-in" `
        -Actual ("Found Addin in Excel.COMAddIns: " + $isRegistered) `
        -Passed $s04Passed -ElapsedMs $sw.Elapsed.TotalMilliseconds
} catch {
    Record-Step -ScenarioId "S04" -Title "First Excel Startup" -Expected "Success" -Actual $_.ToString() -Passed $false -ElapsedMs 0
}

# -----------------------------------------------------------------
# S05: Reference Kutools Coexistence Verification
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $kutoolsPath = "C:\Program Files (x86)\Kutools for Excel"
    $kutoolsExists = Test-Path $kutoolsPath
    $kutoolsFiles = (Get-ChildItem -Path $kutoolsPath -ErrorAction SilentlyContinue).Count
    $sw.Stop()

    $s05Passed = ($kutoolsExists -and $kutoolsFiles -gt 0)
    Record-Step -ScenarioId "S05" -Title "Reference Kutools Coexistence Integrity" `
        -Expected "Kutools directory and binaries remain untouched" `
        -Actual ("Kutools directory exists with " + $kutoolsFiles + " items") `
        -Passed $s05Passed -ElapsedMs $sw.Elapsed.TotalMilliseconds
} catch {
    Record-Step -ScenarioId "S05" -Title "Reference Kutools Coexistence" -Expected "Success" -Actual $_.ToString() -Passed $false -ElapsedMs 0
}

# -----------------------------------------------------------------
# S06: Uninstallation from External Location
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $uninstallScript = Join-Path $DeployBaseDir "uninstall.ps1"
    & powershell.exe -ExecutionPolicy Bypass -File $uninstallScript
    $sw.Stop()

    $regAfter = Get-ItemProperty -Path $addInKey -ErrorAction SilentlyContinue
    $s06Passed = ($regAfter -eq $null)
    Record-Step -ScenarioId "S06" -Title "Clean Uninstallation from External Directory" `
        -Expected "HKCU Addin key completely removed" `
        -Actual ("Registry key exists after uninstall: " + ($regAfter -ne $null)) `
        -Passed $s06Passed -ElapsedMs $sw.Elapsed.TotalMilliseconds
} catch {
    Record-Step -ScenarioId "S06" -Title "Clean Uninstallation" -Expected "Success" -Actual $_.ToString() -Passed $false -ElapsedMs 0
}

# -----------------------------------------------------------------
# S07: Excel Startup After Uninstallation
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $beforePids7 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $afterPids7 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPid7 = $afterPids7 | Where-Object { $beforePids7 -notcontains $_ } | Select-Object -First 1

    $wb = $excel.Workbooks.Add()

    $foundAfterUninstall = $false
    for ($i = 1; $i -le $excel.COMAddIns.Count; $i++) {
        $comItem = $excel.COMAddIns.Item($i)
        if ($comItem.ProgId -eq "ExcelProductivitySuite") {
            $foundAfterUninstall = $true
            break
        }
    }
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if ($newPid7) {
        $proc7 = Get-Process -Id $newPid7 -ErrorAction SilentlyContinue
        if ($proc7) {
            if (-not $proc7.WaitForExit(3000)) { $proc7.Kill(); $proc7.WaitForExit(2000) }
        }
    }
    $sw.Stop()

    $s07Passed = (-not $foundAfterUninstall)
    Record-Step -ScenarioId "S07" -Title "Excel Startup Verification After Uninstall" `
        -Expected "Addin is no longer loaded by Excel" `
        -Actual ("Addin present in Excel.COMAddIns: " + $foundAfterUninstall) `
        -Passed $s07Passed -ElapsedMs $sw.Elapsed.TotalMilliseconds
} catch {
    Record-Step -ScenarioId "S07" -Title "Excel Startup After Uninstall" -Expected "Success" -Actual $_.ToString() -Passed $false -ElapsedMs 0
}

# -----------------------------------------------------------------
# S08: Installation into Custom Target Directory (-TargetDir)
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$customDir = "C:\Users\Yogeswari\AppData\Local\Programs\ExcelProductivitySuite-CustomDir"
try {
    if (Test-Path $customDir) { Remove-Item -Path $customDir -Recurse -Force }
    New-Item -Path $customDir -ItemType Directory -Force | Out-Null
    Copy-Item -Path "$DeployBaseDir\*" -Destination $customDir -Recurse -Force

    $installScript = Join-Path $DeployBaseDir "install.ps1"
    & powershell.exe -ExecutionPolicy Bypass -File $installScript -TargetDir $customDir
    $sw.Stop()

    $regCustom = Get-ItemProperty -Path $addInKey -ErrorAction SilentlyContinue
    $manifestMatchesCustom = $regCustom -ne $null -and $regCustom.Manifest.Contains($customDir.Replace('\', '/'))
    $s08Passed = ($manifestMatchesCustom -and $regCustom.LoadBehavior -eq 3)

    Record-Step -ScenarioId "S08" -Title "Installation into Custom Target Directory (-TargetDir)" `
        -Expected "Manifest points to custom directory and LoadBehavior=3" `
        -Actual ("Manifest: " + $regCustom.Manifest) `
        -Passed $s08Passed -ElapsedMs $sw.Elapsed.TotalMilliseconds
} catch {
    Record-Step -ScenarioId "S08" -Title "Custom Target Directory Installation" -Expected "Success" -Actual $_.ToString() -Passed $false -ElapsedMs 0
}

# -----------------------------------------------------------------
# S09: Recovery from Corrupt Add-in Registration (LoadBehavior = 2)
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    # Simulate Excel disabling addin due to crash (LoadBehavior = 2)
    Set-ItemProperty -Path $addInKey -Name "LoadBehavior" -Value 2 -Type DWord
    $corruptVal = (Get-ItemProperty -Path $addInKey).LoadBehavior

    # Re-run installer to auto-heal registration
    & powershell.exe -ExecutionPolicy Bypass -File $installScript -TargetDir $customDir
    $sw.Stop()

    $healedVal = (Get-ItemProperty -Path $addInKey).LoadBehavior
    $s09Passed = ($corruptVal -eq 2 -and $healedVal -eq 3)

    Record-Step -ScenarioId "S09" -Title "Recovery from Disabled/Corrupt Registration" `
        -Expected "Corrupt LoadBehavior=2 successfully reset to 3 (Load on Startup)" `
        -Actual ("Before: LoadBehavior=" + $corruptVal + " | After Reinstall: LoadBehavior=" + $healedVal) `
        -Passed $s09Passed -ElapsedMs $sw.Elapsed.TotalMilliseconds
} catch {
    Record-Step -ScenarioId "S09" -Title "Recovery from Corrupt Registration" -Expected "Success" -Actual $_.ToString() -Passed $false -ElapsedMs 0
}

# -----------------------------------------------------------------
# S10: Missing Manifest Guard
# -----------------------------------------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$caughtExpectedError = $false
$errMsg = ""
try {
    $nonExistentDir = "C:\NonExistent\Directory\For\Testing"
    $errOutput = & powershell.exe -ExecutionPolicy Bypass -File $installScript -TargetDir $nonExistentDir 2>&1
    if ($LASTEXITCODE -ne 0 -or ($errOutput -match "does not exist on disk|Deployment manifest not found")) {
        $caughtExpectedError = $true
        $errMsg = ($errOutput | Out-String).Trim()
    }
} catch {
    $caughtExpectedError = $true
    $errMsg = $_.ToString().Trim()
}
$sw.Stop()

$s10Passed = $caughtExpectedError
Record-Step -ScenarioId "S10" -Title "Missing Manifest Protective Error Guard" `
    -Expected "Installer halts with actionable error when manifest is missing" `
    -Actual ("Caught expected error: " + $errMsg.Substring(0, [Math]::Min(100, $errMsg.Length)).Replace("`r", " ").Replace("`n", " ")) `
    -Passed $s10Passed -ElapsedMs $sw.Elapsed.TotalMilliseconds

# Clean up temporary test directories
if (Test-Path $customDir) { Remove-Item -Path $customDir -Recurse -Force -ErrorAction SilentlyContinue }
if (Test-Path $DeployBaseDir) { Remove-Item -Path $DeployBaseDir -Recurse -Force -ErrorAction SilentlyContinue }

# Restore registration back to repository build for continued developer workflow
$devInstall = "D:\Kutools-New\installer\install.ps1"
& powershell.exe -ExecutionPolicy Bypass -File $devInstall -Configuration Release | Out-Null

# -----------------------------------------------------------------
# Summary
# -----------------------------------------------------------------
Write-Host "`n==========================================================" -ForegroundColor Cyan
$passedCount = ($results | Where-Object { $_.Passed }).Count
$totalCount = $results.Count
$failedCount = $totalCount - $passedCount
Write-Host ("  LIFECYCLE TEST RESULTS: " + $passedCount + " / " + $totalCount + " PASSED (" + $failedCount + " FAILED)  ") -ForegroundColor $(if ($failedCount -eq 0) { "Green" } else { "Red" })
Write-Host "==========================================================" -ForegroundColor Cyan

# Output JSON
$jsonOut = $results | ConvertTo-Json -Depth 4
Set-Content -Path "D:\Kutools-New\docs\installation-lifecycle-results.json" -Value $jsonOut
Write-Host "Results exported to docs/installation-lifecycle-results.json" -ForegroundColor Gray

if ($failedCount -gt 0) { exit 1 }
exit 0
