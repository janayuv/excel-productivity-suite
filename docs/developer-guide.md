# Excel Productivity Suite — Developer Guide

**Target Environment:** Visual Studio 2022 / .NET Framework 4.8 / Microsoft Office 2016+ (64-bit & 32-bit)  
**Language:** C# (C# latest / 10.0+ with nullable reference types)  

---

## 1. Solution Architecture & Assembly Map

The solution is partitioned into 6 focused projects enforcing strict separation between business logic and host COM interop:

```text
ExcelProductivitySuite.sln
├── src/
│   ├── ExcelProductivitySuite.Core/
│   │   ├── Abstractions/ (ICommandRegistry, IExcelContext, IExcelStateScope)
│   │   ├── Models/       (SuiteCommand, OperationResult, CommandCategory)
│   │   └── Services/     (CommandRegistry)
│   │
│   ├── ExcelProductivitySuite.Engine/
│   │   ├── Text/         (TextExtractor, TextSplitter, CharacterRemover, CaseConverter, WhitespaceNormalizer, AffixAppender)
│   │   ├── Ranges/       (RangeFill, RangeMerger)
│   │   ├── DataCleaning/ (DuplicateFinder, RangeRecordComparer, AdvancedFilter)
│   │   └── Worksheets/   (WorksheetNameValidator, DataSplitter, WorkbookComparer)
│   │
│   ├── ExcelProductivitySuite.UI/
│   │   ├── Themes/       (Brushes.xaml, Typography.xaml, Controls.xaml)
│   │   ├── Controls/     (BaseDialogWindow, LivePreviewTable)
│   │   ├── Models/       (PreviewItem)
│   │   └── Dialogs/      (ToolLauncherWindow, ExtractTextDialog, SplitTextDialog, FindMissingRecordsDialog, MergeRangesDialog, SplitDataToSheetsDialog, AdvancedFilterDialog, CompareWorkbooksDialog, etc.)
│   │
│   ├── ExcelProductivitySuite.Infrastructure/
│   │   └── (Settings, Logging, JSON Serialization)
│   │
│   └── ExcelProductivitySuite.AddIn/
│       ├── ThisAddIn.cs
│       ├── ExcelIntegration/ (ExcelContextAdapter, ExcelStateScope)
│       ├── Ribbon/           (ExcelProductivityRibbon.cs, ExcelProductivityRibbon.xml)
│       └── Startup/          (ToolRegistrar)
│
└── tests/
    └── ExcelProductivitySuite.UnitTests/
        ├── TextOperationsTests.cs
        ├── RangeAndSheetTests.cs
        ├── RecordComparerTests.cs
        ├── RangeMergerTests.cs
        ├── DataSplitterTests.cs
        ├── AdvancedFilterTests.cs
        ├── WorkbookComparerTests.cs
        ├── PerformanceBenchmarkTests.cs
        └── EndToEndIntegrationTests.cs
```

---

## 2. Prerequisites & Building

### Prerequisites
1. **Operating System:** Windows 10 / Windows 11 (64-bit or 32-bit).
2. **Visual Studio:** Visual Studio 2022 Professional or Community with workloads:
   - *.NET desktop development*
   - *Office/SharePoint development* (VSTO tools)
3. **Target Framework:** .NET Framework 4.8 Developer Pack.
4. **Office:** Microsoft Excel 2016, 2019, 2021, or Microsoft 365 desktop.

### Building from Command Line
```powershell
# Restore dependencies
dotnet restore ExcelProductivitySuite.sln

# Build entire solution using Visual Studio MSBuild
& "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" ExcelProductivitySuite.sln /t:Build /p:Configuration=Debug

# Execute automated tests
dotnet test tests/ExcelProductivitySuite.UnitTests/ExcelProductivitySuite.UnitTests.csproj
```

---

## 3. How to Add a New Utility (Step-by-Step)

Adding a new tool takes 4 clean steps without modifying existing features:

### Step 1: Implement Pure Algorithm in `ExcelProductivitySuite.Engine`
Create a pure C# class with zero dependencies on Office COM:
```csharp
namespace ExcelProductivitySuite.Engine.Text
{
    public static class StringReverser
    {
        public static string Reverse(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            char[] chars = input.ToCharArray();
            Array.Reverse(chars);
            return new string(chars);
        }
    }
}
```

### Step 2: Add Unit Tests in `ExcelProductivitySuite.UnitTests`
```csharp
[TestMethod]
public void StringReverser_ReversesCharacters()
{
    Assert.AreEqual("DCBA", StringReverser.Reverse("ABCD"));
}
```

### Step 3: Implement WPF Dialog in `ExcelProductivitySuite.UI`
Create a dialog deriving from `BaseDialogWindow`:
```csharp
public partial class ReverseTextDialog : BaseDialogWindow
{
    private readonly IExcelContext _context;
    public ReverseTextDialog(IExcelContext context)
    {
        _context = context;
        InitializeComponent();
    }
}
```

### Step 4: Register Tool in `ExcelProductivitySuite.AddIn.Startup.ToolRegistrar`
```csharp
registry.RegisterCommand(new SuiteCommand(
    id: "CMD_TEXT_REVERSE",
    title: "Reverse Text",
    description: "Invert character order in selected cells.",
    category: CommandCategory.Text,
    executeAction: ctx =>
    {
        var dlg = new ReverseTextDialog(ctx);
        dlg.ShowModal(ctx);
    },
    keywords: new[] { "reverse", "invert", "flip string" },
    officeIconName: "AutoSummaryToolbar"
));
```
The tool will automatically appear in the **Quick Tool Launcher** (`Ctrl+Shift+P`), searchable and runnable!

---

## 4. Excel COM Automation & Performance Guidelines

1. **Never loop over cells via COM:**
   ```csharp
   // BAD: 10,000 inter-process COM round trips (takes ~35 seconds)
   for (int r = 1; r <= 10000; r++) range.Cells[r, 1].Value2 = ...;

   // GOOD: Read 2D array, process in memory, write back once (takes ~0.08 seconds)
   object[,] values = (object[,])range.Value2;
   // transform values in memory...
   range.Value2 = values;
   ```
2. **Always wrap batch operations in `ExcelStateScope`:**
   ```csharp
   using (new ExcelStateScope(application))
   {
       // ScreenUpdating = false, EnableEvents = false, Calculation = Manual
       // Original settings are 100% guaranteed to be restored in Dispose() finally block.
   }
   ```
3. **STA Thread Rule:** WPF dialogs and COM operations must run on the primary Excel STA UI thread.
4. **Dialog Ownership:** Every modal window must call `ShowModal(IExcelContext context)` to set `WindowInteropHelper.Owner = context.MainWindowHandle`.

---

## 5. Phase 3 Architecture: Data Engines & Abstraction Extensions

### 5.1 Extended `IExcelContext` Abstractions
Phase 3 extended `IExcelContext` to support cross-workbook querying and safe worksheet generation:
```csharp
public interface IExcelContext
{
    // ... Phase 1 & 2 properties ...

    // Phase 3 Extensions
    bool CreateWorksheetWithData(string sheetName, object[,] data);
    IReadOnlyList<string> GetOpenWorkbookNames();
    IReadOnlyList<string> GetWorkbookSheetNames(string workbookName);
    object[,] GetWorkbookRangeValues(string workbookName, string sheetName, string rangeAddress);
}
```

### 5.2 Pure Computational Engines
All Phase 3 data tools are decoupled from Office COM Interop and operate purely on managed 2D matrices:
- `RangeRecordComparer`: Key column indexing, multi-column record composite keys, `SourceOnly`, `TargetOnly`, `Matching`, and `Duplicate` partition matrices.
- `RangeMerger`: Matrix concatenation in `AppendRows` (with optional header deduplication and column alignment) or `SideBySide` dimensions.
- `DataSplitter`: Categorical partitioning with string sanitization (`\ / ? * : [ ]`), Excel 31-character truncation, collision deduplication (`Dept_1`, `Dept_2`), and header repetition.
- `AdvancedFilter`: Multi-rule boolean evaluation (`RuleConnector.And` / `RuleConnector.Or`), supporting `Equals`, `Contains`, `StartsWith`, `EndsWith`, `GreaterThan`, `LessThan`, `IsEmpty`, `IsNotEmpty`, and regex matching.
- `WorkbookComparer`: Cross-workbook structural and cell-level difference engine that leaves source workbooks 100% read-only and outputs a summary matrix.

---

## 6. Phase 4 Architecture: Real Excel Testing, COM Hardening & Packaging

### 6.1 Real Excel Automated Integration Test Runner (`Run-ExcelIntegrationTests.ps1`)
The suite includes a fully automated integration test script (`tests/Run-ExcelIntegrationTests.ps1`) designed to run against active Microsoft Excel 2016+ desktop installations:
- **Out-of-Process Automation:** Instantiates `Excel.Application` via COM, creates disposable workbooks, runs 18 realistic user scenarios, and checks cell-by-cell outcomes.
- **Dynamic Array Marshalling:** Solves PowerShell COM dynamic parameter limitations on 2D object arrays by leveraging an embedded C# helper class (`ComInterop.SetValue2`) using C# `dynamic` late-binding dispatch.
- **Guaranteed Cleanup:** Employs `Marshal.FinalReleaseComObject` across all worksheets, workbooks, and application references within a `finally` block, followed by two cycles of garbage collection to ensure zero orphaned background `EXCEL.EXE` processes.
- **Automated Metric Capture:** Measures separate timings for data synthesis, COM reading, in-memory C# transformation, and COM writing, outputting results in JSON to `docs/real-excel-test-results.json`.

```powershell
# Run real Excel automated integration tests
powershell -ExecutionPolicy Bypass -File tests/Run-ExcelIntegrationTests.ps1
```

### 6.2 Embedded Ribbon XML Architecture
To eliminate all hard-coded development filesystem paths (e.g. `d:\Kutools-New\...`), `ExcelProductivityRibbon.xml` is embedded directly into the compiled `ExcelProductivitySuite.AddIn.dll` as an assembly manifest resource:
```xml
<!-- In ExcelProductivitySuite.AddIn.csproj -->
<ItemGroup>
  <EmbeddedResource Include="Ribbon\ExcelProductivityRibbon.xml" />
</ItemGroup>
```
In `ExcelProductivityRibbon.cs`:
```csharp
public string GetCustomUI(string ribbonID)
{
    var assembly = Assembly.GetExecutingAssembly();
    var resourceName = "ExcelProductivitySuite.AddIn.Ribbon.ExcelProductivityRibbon.xml";
    using (var stream = assembly.GetManifestResourceStream(resourceName))
    {
        if (stream != null)
        {
            using (var reader = new StreamReader(stream))
                return reader.ReadToEnd();
        }
    }
    // Fallback to local file for development reload
    ...
}
```

### 6.3 Resilient COM State Management (`ExcelStateScope`)
`ExcelStateScope.cs` encapsulates temporary suppression of `ScreenUpdating`, `DisplayAlerts`, `EnableEvents`, and `Calculation`. In Phase 4, the disposal pattern was hardened:
- Redundant COM property getters (e.g. `_app.ScreenUpdating != _priorScreenUpdating`) were removed to avoid unnecessary inter-process queries.
- Each property is restored inside its own isolated `try/catch` block, preventing an error on one property (e.g., during protected workbook state) from aborting the restoration of the remaining properties.

### 6.4 Standalone Commercial Distribution Packaging (`package.ps1`)
The automated packaging script (`installer/package.ps1`) builds a clean Release distribution artifact ready for enterprise deployment:
1. Compiles the entire solution using Visual Studio MSBuild in `Release` configuration.
2. Stages the add-in binaries, dependencies, VSTO manifest, and deployment scripts into `dist/ExcelProductivitySuite-v4.0.0-win-x64/`.
3. Verifies binary integrity, ensures all 18 tool assemblies are present, and confirms zero debug symbols or temporary files.
4. Generates a compressed distribution zip: `dist/ExcelProductivitySuite-v4.0.0-win-x64.zip`.
5. Computes and generates cryptographic SHA-256 checksums in `dist/SHA256SUMS.txt`.

```powershell
# Generate release distribution archive
powershell -ExecutionPolicy Bypass -File installer/package.ps1
```

---

## 7. Phase 5 Release Package Validation & Code Signing Guide

### 7.1 Automated Installation Lifecycle Test Suite (`Test-ReleaseInstallationLifecycle.ps1`)
To ensure that production distributions run completely free of the development repository, `tests/Test-ReleaseInstallationLifecycle.ps1` runs 10 end-to-end installation scenarios:
1. **Extraction:** Extracts `ExcelProductivitySuite-v4.0.0-win-x64.zip` into an external location outside `d:\Kutools-New` (`%LOCALAPPDATA%\Programs\ExcelProductivitySuite-ReleaseTest`).
2. **Fresh Installation:** Executes `install.ps1` and validates that the registered manifest URL points directly to the external directory with `|vstolocal` and `LoadBehavior = 3`.
3. **Diagnostics:** Runs `diagnostics.ps1` from the external directory and asserts healthy registration.
4. **Excel Host Load:** Launches real Microsoft Excel via COM and confirms `ExcelProductivitySuite` is enumerated in `Application.COMAddIns`.
5. **Kutools Coexistence:** Confirms all 38 files in `C:\Program Files (x86)\Kutools for Excel` remain undisturbed.
6. **Clean Uninstall:** Runs `uninstall.ps1` and validates complete removal of HKCU registration keys.
7. **Excel Post-Uninstall Launch:** Confirms Excel starts cleanly and no longer loads the suite.
8. **Custom Target Directory:** Tests `-TargetDir` parameter to install into an arbitrary target path.
9. **Corrupt Registration Recovery:** Simulates crash disablement (`LoadBehavior = 2`) and verifies the installer auto-heals it to `3`.
10. **Missing Manifest Guard:** Asserts that providing a missing or invalid `-TargetDir` produces a clean, actionable error rather than a corrupt registry entry.

```powershell
# Run release package installation lifecycle tests
powershell -ExecutionPolicy Bypass -File tests/Test-ReleaseInstallationLifecycle.ps1
```

### 7.2 Enterprise Code Signing Procedure
The release package manifests are signed with an internal developer test certificate (`Subject: CN=ExcelProductivitySuite`, `Thumbprint: 74F1C194A7C449209F1795774588A5D07C4375D6`). For production enterprise rollout where Group Policy requires trusted Authenticode signatures:

1. **Obtain Authenticode Certificate:** Obtain an EV or OV Code Signing certificate from a recognized commercial CA (e.g. DigiCert, Sectigo, GlobalSign) or internal Enterprise Root PKI.
2. **Sign Assemblies:**
   ```powershell
   & "C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x64\signtool.exe" sign /fd SHA256 /a /tr http://timestamp.digicert.com /td SHA256 /sha1 <CertThumbprint> ExcelProductivitySuite.*.dll
   ```
3. **Sign Application Manifest:**
   ```powershell
   & "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\mage.exe" -Sign ExcelProductivitySuite.AddIn.dll.manifest -CertHash <CertThumbprint> -TimestampUrl http://timestamp.digicert.com
   ```
4. **Update and Sign Deployment Manifest:**
   ```powershell
   & "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\mage.exe" -Update ExcelProductivitySuite.AddIn.vsto -AppManifest ExcelProductivitySuite.AddIn.dll.manifest -CertHash <CertThumbprint> -TimestampUrl http://timestamp.digicert.com
   ```
5. **Re-generate Package & Checksums:**
   ```powershell
   powershell -ExecutionPolicy Bypass -File installer/package.ps1
   ```
