# Excel Productivity Suite

[![Build Status](https://img.shields.io/badge/Build-Passing-brightgreen.svg)]()
[![Tests](https://img.shields.io/badge/Tests-55%20Passed-brightgreen.svg)]()
[![Target](https://img.shields.io/badge/Excel-2016%20%7C%202019%20%7C%202021%20%7C%20365-blue.svg)]()
[![Architecture](https://img.shields.io/badge/Platform-Windows%20x64%20%7C%20x86-blue.svg)]()
[![Framework](https://img.shields.io/badge/.NET-4.8%20%2B%20WPF-purple.svg)]()

A Windows desktop Microsoft Excel add-in delivering an Excel productivity experience inspired by the breadth and polish of commercial tools. Built with a custom Office Ribbon, extensive data cleaning and text transformation utilities, WPF dialogs with live before-and-after previews, and high-performance in-memory batch operations.

---

## Highlights

- **Custom Office Ribbon:** Coherent, beautifully organized tabs (`Excel Productivity` and `Workbook & Sheets`) integrating seamlessly into Microsoft Excel.
- **Searchable Quick Tool Launcher (SuperSearch):** Instant keyword search, category filtering, recent tools, favorites, and direct execution (`Ctrl+Shift+P`).
- **High-Performance COM Engine:** Reads entire ranges into managed 2D object arrays, transforms data in memory, and writes results in single-shot batch assignments with automated `ExcelStateScope` restoration.
- **Modern WPF Visual Design System:** Dialogs with Office-aligned color tokens, live before-and-after preview grids, range selection inputs, and keyboard accessibility.
- **Data Integrity & Safety:** Non-destructive defaults, Excel formula preservation, pre-execution duplicate counts, and worksheet naming validation.
- **100% Local & Telemetry-Free:** Operates entirely offline with zero external cloud dependencies.
- **Side-by-Side Coexistence:** Fully verified and tested to coexist alongside existing Kutools installations with zero interference.

---

## Implemented Productivity Utilities (18 Tools Across Phases 1, 2 & 3)

| Command ID | Title | Category | Description |
| :--- | :--- | :--- | :--- |
| `CMD_LAUNCHER` | **Tool Launcher** | General | Searchable tool launcher with instant keyword matching, recents, and favorites filters. |
| `CMD_TEXT_EXTRACT` | **Extract Text** | Text | Extracts substrings before, after, or between delimiters with live preview. |
| `CMD_TEXT_SPLIT` | **Split Cells** | Text | Splits delimited cell values across columns or into multiple rows. |
| `CMD_TEXT_REMOVE_CHARS` | **Remove Characters** | Text | Strips numeric, alphabetic, symbols, control codes, or custom characters. |
| `CMD_TEXT_TRIM_SPACES` | **Trim Spaces** | Text | Cleans excess spaces, leading/trailing gaps, and non-breaking spaces (`\u00A0`). |
| `CMD_TEXT_CHANGE_CASE` | **Change Case** | Text | Converts text to UPPERCASE, lowercase, Proper Case, or Sentence case. |
| `CMD_TEXT_ADD_AFFIX` | **Add Text / Affixes** | Text | Prepends prefixes, appends suffixes, or inserts text at exact character positions. |
| `CMD_RANGE_FILL_BLANKS` | **Fill Blank Cells** | Range & Cells | Populates empty cells using values from above, below, left, right, or constants. |
| `CMD_RANGE_COMPARE` | **Compare Ranges** | Range & Cells | Reconciles datasets, side-by-side cell comparison, and missing item detection. |
| `CMD_SELECT_SPECIAL` | **Select Special** | Range & Cells | Locates and inspects cells matching blank, error, formula, or max/min attributes. |
| `CMD_MERGE_RANGES` | **Merge Ranges** | Range & Cells | Combines compatible ranges via row appending or side-by-side juxtaposition with header alignment. |
| `CMD_DATA_REMOVE_DUPLICATES` | **Remove Duplicates** | Data Cleaning | Identifies and clears duplicate rows with header preservation and column checklist. |
| `CMD_FIND_MISSING_RECORDS` | **Find Missing Records** | Data Cleaning | Reconciles two tables via key columns into SourceOnly, TargetOnly, Matching, or Duplicates. |
| `CMD_ADVANCED_FILTER` | **Advanced Filter** | Data Cleaning | Multi-condition rule filtering with boolean AND/OR logic and non-destructive worksheet export. |
| `CMD_FORMULA_TO_VALUE` | **Formulas to Values** | Data Cleaning | Replaces calculated formulas with static values in-place with user confirmation. |
| `CMD_SHEET_VISIBILITY` | **Sheet Visibility** | Worksheets | Bulk manager for Visible, Hidden, and VeryHidden states with unhide-all. |
| `CMD_SHEET_BATCH_RENAME` | **Rename Worksheets** | Worksheets | Renames sheets in bulk with prefix, suffix, and sequence numbering. |
| `CMD_SPLIT_DATA_TO_SHEETS` | **Split Data to Sheets** | Worksheets | Splits a dataset by a categorical column into individual worksheets with name sanitization. |
| `CMD_COMPARE_WORKBOOKS` | **Compare Workbooks** | Worksheets | Cross-workbook structural and cell-level difference analysis with executive audit reports. |
| `CMD_ABOUT` | **About Suite** | General | Displays product version, architecture, diagnostics, and environment stats. |

---

## Solution Structure

```text
ExcelProductivitySuite.sln
├── src/
│   ├── ExcelProductivitySuite.Core/           # Contracts, models, command registry, abstractions
│   ├── ExcelProductivitySuite.Engine/         # Pure C# data manipulation algorithms (100% unit-tested)
│   ├── ExcelProductivitySuite.UI/             # Modern WPF design system, dialogs, live preview controls
│   ├── ExcelProductivitySuite.Infrastructure/ # Settings, diagnostic logging, JSON serialization
│   └── ExcelProductivitySuite.AddIn/          # VSTO add-in, Ribbon XML, COM adapter, ExcelStateScope
├── tests/
│   └── ExcelProductivitySuite.UnitTests/      # Automated unit and integration test suite (55 tests)
├── installer/                                 # Installation, uninstallation, and diagnostic scripts
└── docs/                                      # Reference audit, UI inventory, feature catalog, ADRs, guides
```

---

## Build & Test

### Build Solution
```powershell
dotnet restore ExcelProductivitySuite.sln
& "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" ExcelProductivitySuite.sln /t:Build /p:Configuration=Debug
```

### Run Automated Tests
```powershell
dotnet test tests/ExcelProductivitySuite.UnitTests/ExcelProductivitySuite.UnitTests.csproj
```
*Current test suite: **55 passed, 0 failed, 100% passing rate** (including 100K-cell performance benchmarks).*

### Register Add-in in Excel
```powershell
powershell -ExecutionPolicy Bypass -File installer/install.ps1
```

### Run Diagnostic Verification
```powershell
powershell -ExecutionPolicy Bypass -File installer/diagnostics.ps1
```

---

## Documentation

- [Production-Readiness Audit](docs/production-readiness-audit.md): Complete audit of data integrity, safeguards, COM lifecycle, benchmarks, and remediations.
- [UI & Accessibility Audit](docs/ui-audit.md): Evaluation of Fluent design tokens, high-DPI scaling, modality, accessibility, and interactive previews.
- [Reference Audit](docs/reference-audit.md): Complete audit of installed Kutools reference files, metadata, and 359 cataloged features.
- [UI Inventory & Design System](docs/ui-inventory.md): Ribbon layout, design tokens, color palette, and control specifications.
- [Feature Catalog](docs/feature-catalog.md): Detailed specifications for each tool across all implementation phases.
- [Architecture Decisions Record (ADR)](docs/architecture-decisions.md): Technical decisions covering COM safety, matrix batching, WPF dialog parenting, and registry.
- [User Guide](docs/user-guide.md): End-user walkthroughs, tips, and feature guides.
- [Developer Guide](docs/developer-guide.md): Architecture map, step-by-step guide to adding new utilities, and guidelines.
