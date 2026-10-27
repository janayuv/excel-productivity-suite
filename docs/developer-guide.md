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
│   │   ├── Ranges/       (RangeFill)
│   │   ├── DataCleaning/ (DuplicateFinder)
│   │   └── Worksheets/   (WorksheetNameValidator)
│   │
│   ├── ExcelProductivitySuite.UI/
│   │   ├── Themes/       (Brushes.xaml, Typography.xaml, Controls.xaml)
│   │   ├── Controls/     (BaseDialogWindow, LivePreviewTable)
│   │   ├── Models/       (PreviewItem)
│   │   └── Dialogs/      (ToolLauncherWindow, ExtractTextDialog, SplitTextDialog, RemoveCharactersDialog, etc.)
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
