# UI Inventory & Visual Design System Specification

**Product Identity:** Excel Productivity Suite  
**UI Technology:** Windows Presentation Foundation (WPF) & Office Ribbon XML  
**Target Environment:** Microsoft Excel 2016 / 2019 / 2021 / Office 365 (Windows Desktop)  

---

## 1. Overview & UI Architecture

Excel Productivity Suite integrates with Excel through two tightly coordinated layers:
1. **The Excel Ribbon (Host Surface):** Custom Office Ribbon XML tabs, groups, large/small command buttons, dropdown menus, and gallery controls embedded directly into the Excel shell.
2. **The Modern WPF Application Layer (Modal & Modeless Dialogs, Task Panes):** A custom WPF design system that presents rich interactive dialogs, live before-and-after data previews, range pickers, progress monitors, and a searchable command launcher.

---

## 2. Ribbon Topology & Group Taxonomy

Based on our empirical audit of the reference application's 359 catalogued tools and user productivity ergonomics, our Ribbon layout is organized into clean, focused tabs:

### 2.1 Primary Tab: `Excel Productivity` (`tabExcelProductivity`)
Focused on active sheet data transformation, text processing, range operations, and formulas:

| Group | Key Commands | Control Types |
| :--- | :--- | :--- |
| **Suite & Search** | • Tool Launcher (SuperSearch)<br>• Home Experience<br>• Settings & Preferences | Large Button with search icon, Menu buttons |
| **Text Tools** | • Add Text (Prefix/Suffix/Position)<br>• Extract Text (Before/After/Between)<br>• Remove Characters (Numeric/Alpha/Symbols)<br>• Trim & Normalize Spaces<br>• Change Case (Upper/Lower/Title/Sentence)<br>• Split Text to Columns/Rows | Dropdown with large split buttons, live dialog launchers |
| **Range & Cells** | • Compare Ranges (Identical/Mismatched)<br>• Swap Ranges<br>• Fill Blank Cells (Values/Linear/Above)<br>• Transform Range (Matrix to Vector)<br>• Round without Formula<br>• Clear Formatting / Formulas | Group with menu buttons and dialog launchers |
| **Merge & Split** | • Combine Rows/Columns without data loss<br>• Split Cells into Columns/Rows<br>• Merge Same Cells<br>• Unmerge Cells & Fill Value | Large split buttons with icon badges |
| **Data Cleaning** | • Remove Duplicate Rows/Values<br>• Convert Formulas to Values<br>• Force Values (Text to Number, Number to Text)<br>• Change Sign of Values<br>• Delete Blank Rows/Columns | Action buttons with confirmation safeguards |
| **Formulas** | • Exact Formula Copy<br>• Convert Reference Styles (Absolute/Relative)<br>• Formula Auditing & Errors | Large button with dropdown menu |

### 2.2 Secondary Tab: `Workbook & Sheets` (`tabWorkbookAndSheets`)
Focused on multi-worksheet management, batch operations, and workbook structure:

| Group | Key Commands | Control Types |
| :--- | :--- | :--- |
| **Worksheet Manager** | • Rename Worksheets in Bulk<br>• Manage Sheet Visibility (Unhide All, Selective)<br>• Create Table of Contents (Index Sheet)<br>• Sort Worksheets (Alpha/Numeric/Color)<br>• Delete Blank / Hidden Sheets | Large primary button and dropdown list |
| **Consolidation** | • Combine Worksheets from Workbook<br>• Split Workbook into Multiple Files<br>• Merge Tables by Key Column | Wizard dialog launchers |
| **Import & Export** | • Export Range to File (CSV, JSON, XML, HTML)<br>• Export Range as Image (PNG, JPG)<br>• Insert File Names / File List | File picker dialog launchers |
| **Utilities & About** | • Productivity Statistics<br>• Diagnostic Logger<br>• About Excel Productivity Suite | Standard utility buttons |

---

## 3. Reusable WPF Design System & Tokens

To deliver a commercial-grade, polished Windows desktop aesthetic, the WPF UI layer uses a centralized design system:

### 3.1 Color Palette & Semantic Tokens
- **Primary Brand (Excel Green Alignment):**
  - `BrandPrimary`: `#107C41` (Office Fluent Green)
  - `BrandHover`: `#0E6B37`
  - `BrandPressed`: `#0B562C`
  - `BrandLight`: `#E6F2EB`
  - `BrandBorder`: `#A1D4B7`
- **Neutral Surface Palette (Light Theme):**
  - `SurfaceBackground`: `#FFFFFF`
  - `SurfaceSecondary`: `#F8F9FA`
  - `SurfaceBorder`: `#E1E3E5`
  - `SurfaceDivider`: `#EAECEF`
  - `TextPrimary`: `#202124`
  - `TextSecondary`: `#5F6368`
  - `TextDisabled`: `#9AA0A6`
- **Feedback & Accent Tokens:**
  - `AccentBlue`: `#0078D4`
  - `StatusSuccess`: `#107C41`
  - `StatusWarning`: `#D83B01`
  - `StatusError`: `#A80000`
  - `StatusInfo`: `#005A9E`

### 3.2 Typography & Sizing
- **Font Family:** `Segoe UI`, `Segoe UI Variable Display`, system fallback sans-serif.
- **Scale:**
  - Window Title / Dialog Header: `16pt`, SemiBold (`#202124`)
  - Section Header: `12pt`, SemiBold (`#323130`)
  - Body Text / Control Labels: `10pt` (`13px`), Regular (`#202124`)
  - Captions / Supertips: `9pt` (`12px`), Regular (`#5F6368`)
  - Monospace Data / Formulas: `Consolas`, `Cascadia Mono`, `10pt`

### 3.3 Reusable Control Standards
1. **BaseDialogWindow:**
   - Standard window chrome with modern custom title bar, icon, title, Minimize/Close.
   - Fixed or resizable with MinWidth/MinHeight constraints.
   - Built-in `OK`, `Apply`, `Cancel`, and `Help` button dock.
   - Built-in notification banner for inline warnings/errors.
2. **RangeSelectorControl:**
   - Textbox with Excel range reference (e.g., `$A$1:$D$100`).
   - Integrated range picker button that minimizes or docks the dialog to allow selecting cells on the worksheet.
   - Real-time cell count and validation badge.
3. **LivePreviewTable:**
   - DataGrid displaying sample rows (up to 10 rows) with "Original" and "Result" columns.
   - Highlights changes dynamically as users adjust settings (delimiters, case options, prefixes).
4. **OperationProgressBar:**
   - Indeterminate or determinate progress bar with item counts (e.g., "Processing cell 4,500 of 10,000...").
   - Cancellation token support allowing clean aborts at any point.

---

## 4. Searchable Tool Launcher (SuperSearch)

The tool launcher is a core differentiator, solving the cognitive load of navigating extensive menus:
- **Shortcut Trigger:** `Ctrl + Shift + P` or Ribbon click.
- **Instant Search:** Fuzzy matching across tool name, category, and alias tags (e.g. searching "case" finds "Change Text Case", "Capitalize Words", "UPPERCASE").
- **Categorized View:** Grouped by Text, Range, Cleaning, Sheets, Workbook, Formulas.
- **Favorites & Recents:** Local user pinning with persistence in `%APPDATA%\ExcelProductivitySuite\user-settings.json`.
- **Keyboard Ergonomics:** `Up`/`Down` arrow navigation, `Enter` to launch, `Esc` to dismiss.
