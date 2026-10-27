# Architecture Decisions Record (ADR)

**Product:** Excel Productivity Suite  
**Platform:** Windows Desktop / Microsoft Excel Desktop (2016+ 64-bit & 32-bit)  
**Development Framework:** .NET Framework 4.8 / C# 9.0+ / WPF / VSTO  

---

## ADR 1: Modular Multi-Project Solution Architecture

### Context
An enterprise-grade Excel add-in requires clear separation of concerns between Excel COM automation, pure data manipulation algorithms, user interface presentation, and infrastructure services. Tightly coupling UI dialogs or string algorithms directly to Excel COM APIs prevents unit testing, degrades maintainability, and increases the risk of COM reference leaks.

### Decision
We partition the solution into clean, focused assemblies:
1. `ExcelProductivitySuite.Core`: Contracts (`ICommandRegistry`, `IExcelContext`, `IProgressNotifier`), domain models, validation primitives, and result types. Zero Office or WPF dependencies.
2. `ExcelProductivitySuite.Engine`: Pure C# data manipulation algorithms (Text extraction, regex, splitting, whitespace normalization, casing, fill algorithms, duplicate detection). 100% unit-testable in memory without Excel.
3. `ExcelProductivitySuite.UI`: WPF design system, resource dictionaries, reusable controls (`RangeSelector`, `LivePreviewTable`), ViewModels, and dialog views.
4. `ExcelProductivitySuite.Infrastructure`: Local settings management, JSON serialization, diagnostic file logging, and telemetry-free error reporting.
5. `ExcelProductivitySuite.AddIn`: VSTO Excel Add-in entry point (`ThisAddIn.cs`), Office Ribbon XML (`Ribbon.xml`), Excel COM adapters, and Win32 window-handle parenting.
6. `ExcelProductivitySuite.UnitTests`: Automated tests for all engine algorithms and core services.

### Consequences
- **Positive:** Algorithms can be exhaustively unit-tested at high speed without starting Excel.
- **Positive:** Clean dependency graph prevents accidental circular references.
- **Positive:** COM objects are confined strictly to the `AddIn` adapter layer.

---

## ADR 2: Safe High-Performance Excel COM Automation Strategy

### Context
Performing cell-by-cell COM calls across large Excel ranges (e.g., 10,000 cells) induces catastrophic performance degradation (often 30+ seconds due to inter-process marshaling overhead). Furthermore, unhandled exceptions during automation can leave Excel in a frozen state if `ScreenUpdating` or `Calculation` remain disabled.

### Decision
1. **Bulk Matrix Marshaling:** For all range operations, read the entire target range into a managed 2D object array (`object[,] values = range.Value2 as object[,]`).
2. **In-Memory Transformations:** The `Engine` transforms the 2D array in managed memory.
3. **Single-Shot Writeback:** Write the updated 2D array back to Excel in one assignment (`range.Value2 = values`).
4. **Guaranteed State Restoration via `ExcelStateScope`:**
   Implement an `IDisposable` state scope:
   ```csharp
   using (new ExcelStateScope(application, disableScreenUpdating: true, disableEvents: true, manualCalculation: true))
   {
       // Perform operations safely
   }
   ```
   The `Dispose()` method strictly restores prior settings inside a `finally` block even if exceptions occur.
5. **COM Reference Management:** Only retain references on the main UI STA thread; avoid passing raw COM interface pointers to background worker threads.

---

## ADR 3: Modern WPF Dialog Integration within Excel Host

### Context
Win32 Office add-ins host modal dialogs. If WPF dialogs are opened without proper owner HWND linking, they can disappear behind the Excel window, cause focus trapping bugs, or fail to respond properly to Alt+Tab.

### Decision
1. **WindowInteropHelper Ownership:** Every WPF dialog created by the add-in sets its `Owner` handle to Excel's main window handle:
   ```csharp
   var helper = new WindowInteropHelper(dialog);
   helper.Owner = new IntPtr(application.Hwnd);
   ```
2. **DPI Awareness:** WPF windows declare Per-Monitor V2 DPI awareness to ensure crisp rendering on modern high-DPI displays and mixed-scaling multi-monitor setups.
3. **Consistent Dialog Base:** A reusable `BaseDialogWindow` provides standard title bar styling, status notifications, validation summaries, and standard action buttons (`OK`, `Cancel`, `Apply`, `Help`).

---

## ADR 4: Centralized Command Registry & Searchable Launcher

### Context
With dozens of utilities, maintaining hard-coded Ribbon callbacks and separate discovery mechanisms leads to code duplication and difficult navigation for users.

### Decision
1. Implement an `ICommandRegistry` service where every suite tool registers its metadata:
   - Unique ID, Name, Category, Tooltip, Keyboard Shortcut, Keywords.
   - Execution action taking the current `IExcelContext`.
   - Availability predicate (e.g. requires active workbook, requires non-empty range).
2. The Ribbon XML invokes commands through a centralized dispatcher method:
   `OnRibbonCommandExecuted(IRibbonControl control)`.
3. The Searchable Tool Launcher queries the registry dynamically, enabling instant filtering, keyword matching, and execution without duplicating tool definitions.

---

## ADR 5: Data Safety & Safeguards for Destructive Operations

### Context
Excel operations that overwrite values, delete rows, or convert formulas can permanently alter user workbooks. Excel's native programmatic undo buffer is cleared when COM operations execute.

### Decision
1. **Non-Destructive Defaults:** Tools default to safe operations (e.g., skip formula cells by default, write results to adjacent columns rather than overwriting).
2. **Live Previews:** Destructive tools provide an interactive preview showing before-and-after rows before the user clicks "Apply" or "OK".
3. **Pre-Execution Validation:** Detect merged cells, protected worksheets, and filtered/hidden ranges before attempting writes. Provide clear, polite warning messages instead of silent failures or crashes.
4. **Operation Summary:** Upon completion of batch operations, display an informative status message (e.g. "Successfully processed 1,248 cells in 0.12 seconds. 14 formula cells were skipped.").
