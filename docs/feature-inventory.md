# Excel Productivity Suite — Comprehensive Feature Inventory & Roadmap

**Solution:** `ExcelProductivitySuite.sln`  
**Host Application:** Microsoft Excel 2016 / 2019 / 2021 / Microsoft 365 (32-bit & 64-bit Windows Desktop)  
**Framework:** .NET Framework 4.8 & Visual Studio Tools for Office (VSTO)  
**UI Subsystem:** Windows Presentation Foundation (WPF) / Modern Fluent Office Palette  
**Version:** 4.5.0 (Batch 5 Worksheet & Workbook Management Milestone)  
**Last Audit Date:** 2026-10-28  

---

## 1. Executive Summary & Inventory Overview

Following the mandatory codebase audit, the **Excel Productivity Suite** feature taxonomy has been unified into a standardized 12-module taxonomy (Modules A through L) aligned with premier Windows desktop Excel productivity toolboxes.

The current implementation comprises:
- **Total Cataloged Capabilities:** 200+ features planned across 12 modules
- **Active & Verified Tools:** 47 commands (45 core productivity utilities + 2 suite navigation/information utilities)
- **Unit Test Coverage:** 153 passing automated unit & performance tests (100% pass rate in 1.0s)
- **Real Excel COM Integration Tests:** 18 passing out-of-process Excel 2016+ scenarios (100% pass rate in 1.25s)
- **Searchable Quick Tool Launcher (SuperSearch):** Integrated multi-category search, active/roadmap status filtering, module badges, and keyword scoring (`Ctrl+Shift+P` / KeyTip: `TL`)

---

## 2. Taxonomy Status Classification

Every feature in the inventory is assigned one of the five mandatory status classifications:
1. **Implemented and verified:** Code written, UI integrated, unit and real Excel integration tests passing.
2. **Implemented but needs verification:** Code exists, requires comprehensive edge-case or performance verification.
3. **Partially implemented:** Backend engine exists, but UI or command wiring is incomplete.
4. **Planned:** Fully specified in the roadmap, scheduled for controlled implementation batch.
5. **Blocked:** Deferred or blocked due to COM Interop, Excel 2016 constraints, or security limitations.

---

## 3. Discrepancy Analysis (Audited vs. Previous Documentation)

An independent code inspection of `ExcelProductivitySuite.sln`, `docs/feature-catalog.md`, and the Windows Registry identified three key discrepancies in the historical documentation:

| Documented Claim | Audited Reality | Remediation in Batch 1 |
| :--- | :--- | :--- |
| `feature-catalog.md:137` listed `CMD_RANGE_FIND_REPLACE` as "Implemented" | No engine class, command registration, or UI dialog existed in the repository. | Classified as **Planned** in Module C (Batch 4). Removed false implementation claim. |
| `feature-catalog.md:191` listed `CMD_BATCH_CELL_FORMAT` as "Implemented" | No engine class, command registration, or UI dialog existed in the repository. | Classified as **Planned** in Module H (Batch 7). Removed false implementation claim. |
| `feature-catalog.md:124` listed `CMD_TEXT_INSERT_AT_POS` as a separate command | Implemented directly inside `CMD_TEXT_ADD_AFFIX` (`AddTextDialog` / `PositionInserter`). | Consolidated into `CMD_TEXT_ADD_AFFIX` with explicit position insertion support. |
| Dual Registry in HKCU for development path | `ExcelProductivitySuite.AddIn` pointed to `bin/Debug` while `ExcelProductivitySuite` pointed to `dist`. | Standardized installer to register both keys cleanly to the target production deployment path. |

---

## 4. Complete Active Feature Inventory (27 Implemented & Verified Tools)

| ID | Title | Module | Category | Engine Class | UI Dialog / Action | Test Mapping | Status |
| :--- | :--- | :---: | :--- | :--- | :--- | :--- | :--- |
| `CMD_LAUNCHER` | **Tool Launcher** | L | Navigation & Productivity | `CommandRegistry.cs` | `ToolLauncherWindow.xaml` | `CommandRegistryTests`, `EndToEndIntegrationTests` | Implemented and verified |
| `CMD_ABOUT` | **About Suite** | General | General / Info | System Diagnostics | `AboutDialog.xaml` | `ThisAddIn` | Implemented and verified |
| `CMD_TEXT_EXTRACT` | **Extract Text** | A | Text Tools | `TextExtractor.cs` | `ExtractTextDialog.xaml` | `TextOperationsTests`, `Benchmark_TextExtraction` | Implemented and verified |
| `CMD_TEXT_SPLIT` | **Split Cells** | A | Text Tools | `TextSplitter.cs` | `SplitTextDialog.xaml` | `TextOperationsTests`, `EdgeCaseAndRegressionTests` | Implemented and verified |
| `CMD_TEXT_REMOVE_CHARS` | **Remove Characters** | A | Text Tools | `CharacterRemover.cs` | `RemoveCharactersDialog.xaml` | `TextOperationsTests` | Implemented and verified |
| `CMD_TEXT_TRIM_SPACES` | **Trim Spaces** | A | Text Tools | `WhitespaceNormalizer.cs` | `WhitespaceDialog.xaml` | `TextOperationsTests`, `EdgeCaseAndRegressionTests` | Implemented and verified |
| `CMD_TEXT_CHANGE_CASE` | **Change Case** | A | Text Tools | `CaseConverter.cs` | `ChangeCaseDialog.xaml` | `TextOperationsTests` | Implemented and verified |
| `CMD_TEXT_ADD_AFFIX` | **Add Text / Affixes** | A | Text Tools | `AffixAppender.cs` | `AddTextDialog.xaml` | `TextOperationsTests`, `EdgeCaseAndRegressionTests` | Implemented and verified |
| `CMD_TEXT_COMBINE` | **Combine Cells** | A | Text Tools | `TextCombiner.cs` | `CombineCellsDialog.xaml` | `TextOperationsTests`, `Benchmark_TextCombiner` | Implemented and verified |
| `CMD_TEXT_REVERSE` | **Reverse Text** | A | Text Tools | `TextReverser.cs` | `ReverseTextDialog.xaml` | `TextOperationsTests`, `Benchmark_TextReverser` | Implemented and verified |
| `CMD_TEXT_REMOVE_BY_POS` | **Remove by Position** | A | Text Tools | `PositionRemover.cs` | `RemoveByPositionDialog.xaml` | `TextOperationsTests` | Implemented and verified |
| `CMD_TEXT_CLEAN_NONPRINTING` | **Clean Characters** | A | Text Tools | `TextCleaner.cs` | `CleanTextDialog.xaml` | `TextOperationsTests` | Implemented and verified |
| `CMD_TEXT_CONVERT_NUMBERS` | **Convert to Numbers** | A | Text Tools | `NumberTextConverter.cs` | `ConvertNumbersDialog.xaml` | `TextOperationsTests` | Implemented and verified |
| `CMD_TEXT_CONVERT_DATES` | **Convert Dates** | A | Text Tools | `DateTextConverter.cs` | `ConvertDatesDialog.xaml` | `TextOperationsTests` | Implemented and verified |
| `CMD_TEXT_SPELL_NUMBER` | **Spell Numbers** | A | Text Tools | `NumberSpeller.cs` | `SpellNumberDialog.xaml` | `TextOperationsTests` | Implemented and verified |
| `CMD_RANGE_FILL_BLANKS` | **Fill Blank Cells** | B | Range & Cells | `RangeFill.cs` | `FillBlanksDialog.xaml` | `RangeAndSheetTests`, `Benchmark_FillBlanks` | Implemented and verified |
| `CMD_RANGE_TRANSPOSE` | **Transpose Range** | B | Range & Cells | `RangeTransposer.cs` | `TransposeRangeDialog.xaml` | `RangeAndSheetTests`, `Benchmark_RangeTransposer` | Implemented and verified |
| `CMD_RANGE_FLIP_HORIZONTAL` | **Flip Range** | B | Range & Cells | `RangeFlipper.cs` | `FlipRangeDialog.xaml` | `RangeAndSheetTests` | Implemented and verified |
| `CMD_RANGE_SWAP` | **Swap Ranges** | B | Range & Cells | `RangeSwapper.cs` | `SwapRangesDialog.xaml` | `RangeAndSheetTests` | Implemented and verified |
| `CMD_RANGE_CHANGE_SIGN` | **Change Sign (+/-)** | B | Range & Cells | `SignChanger.cs` | `ChangeSignDialog.xaml` | `RangeAndSheetTests` | Implemented and verified |
| `CMD_RANGE_ROUND_VALUES` | **Round Values** | B | Range & Cells | `ValueRounder.cs` | `RoundValuesDialog.xaml` | `RangeAndSheetTests`, `Benchmark_ValueRounder` | Implemented and verified |
| `CMD_RANGE_RANDOM_SORT` | **Random Sort** | B | Range & Cells | `RangeRandomizer.cs` | `RandomSortDialog.xaml` | `RangeAndSheetTests`, `Benchmark_RangeRandomizer` | Implemented and verified |
| `CMD_RANGE_MERGE_KEEP_DATA` | **Merge (Keep Data)** | B | Range & Cells | `SafeRangeMerger.cs` | `MergeKeepDataDialog.xaml` | `RangeAndSheetTests` | Implemented and verified |
| `CMD_RANGE_INSERT_BLANK_ROWS` | **Insert Blank Rows** | B | Range & Cells | `BlankRowManager.cs` | `InsertBlankRowsDialog.xaml` | `RangeAndSheetTests` | Implemented and verified |
| `CMD_RANGE_DELETE_BLANK_ROWS` | **Delete Blank Rows** | B | Range & Cells | `BlankRowManager.cs` | `DeleteBlankRowsDialog.xaml` | `RangeAndSheetTests` | Implemented and verified |
| `CMD_RANGE_COMPARE` | **Compare Ranges** | C | Find, Select & Compare | `RangeComparer.cs` | `CompareRangesDialog.xaml` | `RangeAndSheetTests`, `Benchmark_RangeComparer` | Implemented and verified |
| `CMD_SELECT_SPECIAL` | **Select Special** | C | Find, Select & Compare | `RangeSelector.cs` | `SelectSpecialDialog.xaml` | `RangeAndSheetTests` | Implemented and verified |
| `CMD_MERGE_RANGES` | **Merge Ranges** | B | Range & Cells | `RangeMerger.cs` | `MergeRangesDialog.xaml` | `RangeMergerTests`, `Benchmark_RangeMerger` | Implemented and verified |
| `CMD_DATA_REMOVE_DUPLICATES` | **Remove Duplicates** | E | Data & Filtering | `DuplicateFinder.cs` | `RemoveDuplicatesDialog.xaml` | `EdgeCaseAndRegressionTests`, `Benchmark_DuplicateDetection` | Implemented and verified |
| `CMD_FIND_MISSING_RECORDS` | **Find Missing Records** | E | Data & Filtering | `RangeRecordComparer.cs` | `FindMissingRecordsDialog.xaml` | `RecordComparerTests`, `Benchmark_RangeRecordComparer` | Implemented and verified |
| `CMD_ADVANCED_FILTER` | **Advanced Filter** | E | Data & Filtering | `AdvancedFilter.cs` | `AdvancedFilterDialog.xaml` | `AdvancedFilterTests`, `Benchmark_AdvancedFilter` | Implemented and verified |
| `CMD_FORMULA_TO_VALUE` | **Formulas to Values** | G | Formulas & Calculation | `ExcelContextAdapter.cs` | In-Place Batch Assignment | `EndToEndIntegrationTests`, Real Excel `T07` | Implemented and verified |
| `CMD_SHEET_VISIBILITY` | **Sheet Visibility** | D | Worksheet & Workbook | `ExcelContextAdapter.cs` | `ManageWorksheetsDialog.xaml` | Real Excel `T11` | Implemented and verified |
| `CMD_SHEET_BATCH_RENAME` | **Rename Worksheets** | D | Worksheet & Workbook | `WorksheetNameValidator.cs` | `BatchRenameWorksheetsDialog.xaml`| `RangeAndSheetTests`, Real Excel `T15` | Implemented and verified |
| `CMD_SPLIT_DATA_TO_SHEETS` | **Split Data to Sheets**| D | Worksheet & Workbook | `DataSplitter.cs` | `SplitDataToSheetsDialog.xaml` | `DataSplitterTests`, `Benchmark_DataSplitter` | Implemented and verified |
| `CMD_COMPARE_WORKBOOKS` | **Compare Workbooks** | D | Worksheet & Workbook | `WorkbookComparer.cs` | `CompareWorkbooksDialog.xaml` | `WorkbookComparerTests` | Implemented and verified |

---

## 5. Comprehensive 12-Module Taxonomy & Planned Roadmap (Modules A - L)

### Module A: Text Tools (Target: 18 Tools)
- [x] `CMD_TEXT_EXTRACT`: Extract substrings before, after, or between delimiters. *(Active)*
- [x] `CMD_TEXT_SPLIT`: Split delimited cell contents into columns or rows. *(Active)*
- [x] `CMD_TEXT_REMOVE_CHARS`: Remove numeric, alphabetic, non-alphanumeric, or custom characters. *(Active)*
- [x] `CMD_TEXT_TRIM_SPACES`: Clean excess, leading/trailing, and non-breaking spaces. *(Active)*
- [x] `CMD_TEXT_CHANGE_CASE`: Convert text to UPPERCASE, lowercase, Proper Case, or Sentence case. *(Active)*
- [x] `CMD_TEXT_ADD_AFFIX`: Prepend prefix, append suffix, or insert at character position. *(Active)*
- [x] `CMD_TEXT_REVERSE`: Reverse cell text by character order (surrogate-safe), words, or custom delimiter. *(Active - Batch 2)*
- [x] `CMD_TEXT_COMBINE`: Concatenate rows, columns, or entire selection with custom delimiters and blank skipping. *(Active - Batch 2)*
- [x] `CMD_TEXT_CLEAN_NONPRINTING`: Strip ASCII control codes (0-31, 127), zero-width characters, and decode HTML entities. *(Active - Batch 2)*
- [x] `CMD_TEXT_CONVERT_NUMBERS`: Convert numbers stored as text, strip currencies, accounting parens, and percentages into true numeric cells. *(Active - Batch 2)*
- [x] `CMD_TEXT_CONVERT_DATES`: Recognize and convert varied text date formats into standardized ISO, EU, US, or long dates. *(Active - Batch 2)*
- [x] `CMD_TEXT_SPELL_NUMBER`: Convert numbers/currencies to written English words up to trillions (SpellNumber). *(Active - Batch 2)*
- [x] `CMD_TEXT_REMOVE_BY_POS`: Remove characters by exact position (left N, right N, middle range, before/after delimiter). *(Active - Batch 2)*
- [ ] `CMD_TEXT_BATCH_REPLACE`: Multi-pattern string and punctuation batch replacement. *(Planned - Batch 2+)*
- [ ] `CMD_TEXT_EXTRACT_NUMBERS`: Extract only numeric digits from mixed text strings. *(Planned - Batch 2+)*
- [ ] `CMD_TEXT_EXTRACT_EMAILS`: Extract email addresses from free-form text. *(Planned - Batch 2+)*
- [ ] `CMD_TEXT_EXTRACT_URLS`: Extract web URLs and domain links into adjacent cells. *(Planned - Batch 2+)*

### Module B: Range and Cell Operations (Target: 18 Tools)
- [x] `CMD_RANGE_FILL_BLANKS`: Populate blank cells from above, below, left, right, or constants. *(Active)*
- [x] `CMD_MERGE_RANGES`: Append rows with header alignment or side-by-side concatenation. *(Active)*
- [x] `CMD_RANGE_TRANSPOSE`: Transpose rows to columns and columns to rows with formatting. *(Active - Batch 3)*
- [x] `CMD_RANGE_FLIP_HORIZONTAL`: Mirror/flip columns horizontally left-to-right. *(Active - Batch 3)*
- [x] `CMD_RANGE_FLIP_VERTICAL`: Mirror/flip rows vertically top-to-bottom. *(Active - Batch 3)*
- [x] `CMD_RANGE_SWAP`: Swap contents and formatting of two non-overlapping selections. *(Active - Batch 3)*
- [x] `CMD_RANGE_CHANGE_SIGN`: Invert numeric signs (+ to - or - to +) in selected range. *(Active - Batch 3)*
- [x] `CMD_RANGE_ROUND_VALUES`: Round numeric cells in-place to precision without adding formulas. *(Active - Batch 3)*
- [x] `CMD_RANGE_RANDOM_SORT`: Shuffle rows or columns in randomized sequence. *(Active - Batch 3)*
- [x] `CMD_RANGE_MERGE_KEEP_DATA`: Merge cells while preserving all cell data using custom separators. *(Active - Batch 3)*
- [x] `CMD_RANGE_INSERT_BLANK_ROWS`: Insert blank rows at regular intervals or value changes. *(Active - Batch 3)*
- [x] `CMD_RANGE_DELETE_BLANK_ROWS`: Delete completely empty rows or columns safely. *(Active - Batch 3)*
- [ ] `CMD_RANGE_COPY_NONCONTIGUOUS`: Copy disjoint multi-area selection into contiguous target. *(Planned - Batch 3+)*
- [ ] `CMD_RANGE_PASTE_VISIBLE_ONLY`: Paste clipboard data strictly into visible cells skipping hidden rows. *(Planned - Batch 3+)*
- [ ] `CMD_RANGE_UNMERGE_FILL`: Unmerge composite cells and fill all constituent cells with value. *(Planned - Batch 3+)*
- [ ] `CMD_RANGE_DECIMALS`: Increase or decrease decimal places in bulk. *(Planned - Batch 3+)*
- [ ] `CMD_RANGE_RANDOM_SAMPLE`: Extract random sample of N records or P% from dataset. *(Planned - Batch 3+)*

### Module C: Find, Select, and Compare (Target: 14 Tools)
- [x] `CMD_RANGE_COMPARE`: Side-by-side cell comparison and discrepancy identification. *(Active)*
- [x] `CMD_SELECT_SPECIAL`: Select blanks, constants, formulas, errors, max/min cells. *(Active)*
- [x] `CMD_FIND_DUPLICATES`: Identify, highlight, or select duplicate or unique values across columns with occurrence rules. *(Active - Batch 4)*
- [x] `CMD_ADVANCED_FIND_REPLACE`: Search and batch replace across multiple selected worksheets or entire workbooks with preview. *(Active - Batch 4)*
- [x] `CMD_SELECT_MAX_MIN`: Locate and select maximum or minimum numeric cells across selection, rows, or columns. *(Active - Batch 4)*
- [x] `CMD_SELECT_ROW_DIFFERENCES`: Compare each cell in a row or column against a baseline reference cell and select mismatches. *(Active - Batch 4)*
- [x] `CMD_FIND_BY_COMMENT`: Select, list, filter, and audit cells containing legacy comments or modern notes across worksheets. *(Active - Batch 4)*
- [ ] `CMD_FIND_BY_FORMAT`: Find cells matching font color, fill color, font weight, or borders. *(Planned - Batch 4+)*
- [ ] `CMD_FIND_BY_VALIDATION`: Locate all cells governed by data validation rules. *(Planned - Batch 4+)*
- [ ] `CMD_SELECT_DUPLICATES`: Quick one-click selection of duplicate cells. *(Planned - Batch 4+)*
- [ ] `CMD_HIGHLIGHT_DUPLICATES`: Direct conditional color highlighting rule generator for duplicate rows. *(Planned - Batch 4+)*
- [ ] `CMD_SELECT_COLUMN_DIFFERENCES`: Direct single-click column difference selector. *(Planned - Batch 4+)*

### Module D: Worksheet and Workbook Management (Target: 18 Tools)
- [x] `CMD_SHEET_VISIBILITY`: Bulk toggle Visible, Hidden, and VeryHidden worksheet states. *(Active)*
- [x] `CMD_SHEET_BATCH_RENAME`: Batch rename sheets with prefix, suffix, and sequence numbers. *(Active)*
- [x] `CMD_SPLIT_DATA_TO_SHEETS`: Distribute rows into separate sheets based on key column. *(Active)*
- [x] `CMD_COMPARE_WORKBOOKS`: Cross-workbook structural and cell difference auditing. *(Active)*
- [x] `CMD_SHEET_COMBINE`: Consolidate multiple selected worksheets into a single master worksheet with header options. *(Active - Batch 5)*
- [x] `CMD_WB_COMBINE_FILES`: Batch import worksheets from multiple external Excel files into one central workbook. *(Active - Batch 5)*
- [x] `CMD_WB_SPLIT_FILES`: Save each worksheet as an independent Excel (.xlsx), PDF, or CSV file. *(Active - Batch 5)*
- [x] `CMD_SHEET_SORT_TABS`: Reorder worksheet tabs alphabetically (A-Z / Z-A), numerically, or reversed. *(Active - Batch 5)*
- [x] `CMD_SHEET_CREATE_INDEX`: Generate Table of Contents sheet with clickable hyperlinks to all tabs. *(Active - Batch 5)*
- [x] `CMD_SHEET_DELETE_BLANKS`: Scan and delete completely empty worksheets after confirmation. *(Active - Batch 5)*
- [ ] `CMD_SHEET_COPY_PAGE_SETUP`: Copy margins, orientation, and scaling across sheets. *(Planned - Batch 5+)*
- [ ] `CMD_SHEET_FREEZE_MULTI`: Synchronize Freeze Panes across all worksheets simultaneously. *(Planned - Batch 5+)*
- [ ] `CMD_SHEET_SEQUENCE_NAMES`: Create a sequence of new worksheets with custom naming pattern. *(Planned - Batch 5+)*
- [ ] `CMD_SHEET_CONSOLIDATE`: Aggregate values across identical sheets using standard math functions. *(Planned - Batch 5+)*
- [ ] `CMD_WB_OPEN_FOLDER`: Open Windows File Explorer at active workbook's containing folder. *(Planned - Batch 5)*
- [ ] `CMD_WB_COPY_PATH`: Copy full workbook file path to clipboard. *(Planned - Batch 5)*
- [ ] `CMD_WB_PROPERTIES`: Inspect detailed workbook properties, sheet count, author, and size. *(Planned - Batch 5)*

### Module E: Data, Filtering, and Statistics (Target: 15 Tools)
- [x] `CMD_DATA_REMOVE_DUPLICATES`: Remove duplicate rows with key-column selection and header preservation. *(Active)*
- [x] `CMD_FIND_MISSING_RECORDS`: Reconcile two tables and classify into SourceOnly, TargetOnly, Matching. *(Active)*
- [x] `CMD_ADVANCED_FILTER`: Multi-condition rule filtering with AND/OR logic and non-destructive export. *(Active)*
- [ ] `CMD_DATA_ADVANCED_SORT`: Multi-level sorting by text length, month/weekday, or color. *(Planned - Batch 6)*
- [ ] `CMD_DATA_FILTER_BY_COLOR`: Filter rows by background fill or font color. *(Planned - Batch 6)*
- [ ] `CMD_DATA_COUNT_BY_COLOR`: Count, sum, and average cells by cell fill or font color. *(Planned - Batch 6)*
- [ ] `CMD_DATA_MERGE_TABLES`: Upsert and update target table records matching secondary lookup keys. *(Planned - Batch 6)*
- [ ] `CMD_DATA_QUICK_SUBTOTAL`: Automatically insert group subtotals and summaries at key changes. *(Planned - Batch 6)*
- [ ] `CMD_DATA_FREQUENCY_STATS`: Calculate occurrence frequency distribution table for column values. *(Planned - Batch 6)*
- [ ] `CMD_DATA_EXTRACT_DISTINCT`: Filter and copy distinct rows to a new destination worksheet. *(Planned - Batch 6)*

### Module F: Data Validation and Dropdown Lists (Target: 12 Tools)
- [ ] `CMD_VALIDATION_CREATE_DROPDOWN`: Create standard Excel in-cell dropdown list from items/range. *(Planned - Batch 6)*
- [ ] `CMD_VALIDATION_SEARCHABLE_DROPDOWN`: Floating searchable popup interface for large validation lists. *(Planned - Batch 6)*
- [ ] `CMD_VALIDATION_MULTI_SELECT`: Allow multi-item selection in dropdown separated by commas. *(Planned - Batch 6)*
- [ ] `CMD_VALIDATION_PREVENT_DUPS`: Apply validation rule that rejects duplicate entries in real-time. *(Planned - Batch 6)*
- [ ] `CMD_VALIDATION_EMAIL_RULE`: Apply regex-based validation rule requiring valid email syntax. *(Planned - Batch 6)*
- [ ] `CMD_VALIDATION_PHONE_RULE`: Apply validation rule enforcing phone number patterns. *(Planned - Batch 6)*
- [ ] `CMD_VALIDATION_DATE_RANGE`: Restrict entry to dates within specified calendar boundaries. *(Planned - Batch 6)*
- [ ] `CMD_VALIDATION_HIGHLIGHT_INVALID`: Scan range and highlight all cells with invalid validation data. *(Planned - Batch 6)*
- [ ] `CMD_VALIDATION_BULK_REMOVE`: Remove data validation rules across selected ranges or sheets. *(Planned - Batch 6)*
- [ ] `CMD_VALIDATION_COPY_RULES`: Copy validation rules from source range to target ranges. *(Planned - Batch 6)*

### Module G: Formula and Calculation Utilities (Target: 14 Tools)
- [x] `CMD_FORMULA_TO_VALUE`: Convert formulas to calculated values in-place with safeguards. *(Active)*
- [ ] `CMD_FORMULA_CONVERT_REF`: Convert cell references between Absolute, Row-Absolute, Col-Absolute, Relative. *(Planned - Batch 7)*
- [ ] `CMD_FORMULA_EXACT_COPY`: Copy formulas to new ranges without shifting relative cell references. *(Planned - Batch 7)*
- [ ] `CMD_FORMULA_AUDIT_PRECEDENTS`: Trace and visually navigate precedent cells contributing to formula. *(Planned - Batch 7)*
- [ ] `CMD_FORMULA_AUDIT_DEPENDENTS`: Trace and visually navigate all dependent cells of active cell. *(Planned - Batch 7)*
- [ ] `CMD_FORMULA_LOCATE_ERRORS`: Find and cycle through formula errors (#N/A, #VALUE!, #REF!, #DIV/0!). *(Planned - Batch 7)*
- [ ] `CMD_FORMULA_BATCH_REPLACE_REF`: Replace sheet names or range addresses within formulas in bulk. *(Planned - Batch 7)*
- [ ] `CMD_FORMULA_MATH_OPERATIONS`: Apply math calculation (+, -, *, /) directly to cell values in selection. *(Planned - Batch 7)*
- [ ] `CMD_FORMULA_EXPLAIN`: Deconstruct and explain nested formulas with deterministic parser. *(Planned - Batch 7)*
- [ ] `CMD_FORMULA_TEMPLATES`: Insert pre-built financial, date, and math formula templates. *(Planned - Batch 7)*

### Module H: Formatting and Design (Target: 14 Tools)
- [ ] `CMD_FORMAT_BATCH_NUMBERS`: Batch apply number, currency, decimal, and percent formats. *(Planned - Batch 7)*
- [ ] `CMD_FORMAT_DATE_STANDARDS`: Standardize date cells to ISO 8601 (YYYY-MM-DD) or custom formats. *(Planned - Batch 7)*
- [ ] `CMD_FORMAT_HIGHLIGHT_CELL_TYPES`: Color-code formulas, constants, locked cells, and errors. *(Planned - Batch 7)*
- [ ] `CMD_FORMAT_BORDERS`: Apply accounting borders, double bottom lines, and header styling. *(Planned - Batch 7)*
- [ ] `CMD_FORMAT_ROW_HEIGHT`: Set uniform row heights or smart auto-fit across multiple tables. *(Planned - Batch 7)*
- [ ] `CMD_FORMAT_COL_WIDTH`: Set uniform column widths or padded auto-fit across selections. *(Planned - Batch 7)*
- [ ] `CMD_PRINT_SET_PRINT_AREAS`: Configure multiple print areas across sheets simultaneously. *(Planned - Batch 7)*
- [ ] `CMD_PRINT_INSERT_PAGE_BREAKS`: Insert horizontal page breaks every N rows or on category changes. *(Planned - Batch 7)*
- [ ] `CMD_PRINT_REPEATING_HEADERS`: Configure rows and columns to repeat on every printed page. *(Planned - Batch 7)*
- [ ] `CMD_PRINT_COPY_PAGE_SETUP`: Copy print orientation, margins, and paper size across sheets. *(Planned - Batch 7)*

### Module I: Import, Export, and File Utilities (Target: 14 Tools)
- [ ] `CMD_FILE_IMPORT_CSV`: Import delimited text and CSV files with custom encodings. *(Planned - Batch 8)*
- [ ] `CMD_FILE_EXPORT_RANGE_CSV`: Export selected range or active table to CSV or JSON file. *(Planned - Batch 8)*
- [ ] `CMD_FILE_EXPORT_SHEETS_PDF`: Batch export selected sheets or workbook to PDF format. *(Planned - Batch 8)*
- [ ] `CMD_FILE_BATCH_CONVERT`: Batch convert folders of .xls files to modern .xlsx format. *(Planned - Batch 8)*
- [ ] `CMD_FILE_FOLDER_LISTING`: Generate file and folder directory inventory table in worksheet. *(Planned - Batch 8)*
- [ ] `CMD_FILE_IMPORT_PICTURES`: Batch import picture files into cells with automated sizing. *(Planned - Batch 8)*
- [ ] `CMD_FILE_MATCH_PICTURES`: Match picture filenames against cell keys (SKUs, IDs) and insert. *(Planned - Batch 8)*
- [ ] `CMD_FILE_EXPORT_CHARTS`: Export embedded charts as high-resolution PNG, JPEG, or SVG images. *(Planned - Batch 8)*
- [ ] `CMD_FILE_INSERT_TEXT_FILE`: Insert raw text from an external file into active cell cursor. *(Planned - Batch 8)*

### Module J: Protection and Recovery (Target: 10 Tools)
- [ ] `CMD_SECURITY_BATCH_PROTECT`: Batch protect multiple worksheets with authorized password. *(Planned - Batch 9)*
- [ ] `CMD_SECURITY_BATCH_UNPROTECT`: Batch unprotect multiple worksheets using authorized password. *(Planned - Batch 9)*
- [ ] `CMD_SECURITY_WORKBOOK_ENCRYPT`: Set or update workbook open / modify password. *(Planned - Batch 9)*
- [ ] `CMD_SECURITY_WINDOW_VISIBILITY`: Hide or unhide workbook windows across multi-document sessions. *(Planned - Batch 9)*
- [ ] `CMD_RECOVERY_CREATE_SNAPSHOT`: Save in-session workbook snapshot for safe trial operations. *(Planned - Batch 9)*
- [ ] `CMD_RECOVERY_RESTORE_SNAPSHOT`: Revert active workbook back to previously captured snapshot. *(Planned - Batch 9)*

### Module K: Charts and Analysis (Target: 10 Tools)
- [ ] `CMD_CHART_TEMPLATES`: Apply reusable corporate chart color themes and clean layouts. *(Planned - Batch 9)*
- [ ] `CMD_CHART_GANTT`: Generate project timeline Gantt charts from start date and duration. *(Planned - Batch 9)*
- [ ] `CMD_CHART_WATERFALL`: Create financial waterfall / bridge charts compatible with Excel 2016. *(Planned - Batch 9)*
- [ ] `CMD_CHART_TARGET_VS_ACTUAL`: Generate target vs. actual variance and bullet charts. *(Planned - Batch 9)*
- [ ] `CMD_ANALYSIS_QUICK_STATS`: Display instant statistical dashboard (Count, Sum, Mean, Median, Min, Max). *(Planned - Batch 9)*
- [ ] `CMD_ANALYSIS_PIVOT_HELPERS`: Batch refresh PivotTables, format number fields, toggle classic layout. *(Planned - Batch 9)*

### Module L: Navigation and User Productivity (Target: 12 Tools)
- [x] `CMD_LAUNCHER`: Searchable Tool Catalog (SuperSearch) with instant keyword matching. *(Active)*
- [x] `CMD_ABOUT`: Product version, environment diagnostics, and build configuration. *(Active)*
- [ ] `CMD_NAV_PANE`: Docked task pane for rapid navigation across sheets, tables, and names. *(Planned - Batch 10)*
- [ ] `CMD_NAV_NAMED_RANGES`: Named range inspector, search, and navigation tool. *(Planned - Batch 10)*
- [ ] `CMD_NAV_COLUMNS`: Column navigator and fast visibility toggler for wide tables. *(Planned - Batch 10)*
- [ ] `CMD_NAV_CELL_PREVIEW`: Dedicated inspection card for viewing long cell contents. *(Planned - Batch 10)*
- [ ] `CMD_NAV_FORMULA_EDITOR`: Multi-line expanded formula editing window with syntax layout. *(Planned - Batch 10)*
- [ ] `CMD_NAV_DATE_PICKER`: Floating calendar date picker for rapid date entry. *(Planned - Batch 10)*
- [ ] `CMD_NAV_GRID_FOCUS`: Row and column crosshair focus highlighting for reading wide datasets. *(Planned - Batch 10)*
- [ ] `CMD_NAV_DATA_ENTRY_FORM`: Dynamic modal data entry form for structured tables. *(Planned - Batch 10)*

---

## 6. Implementation Batch Plan (Batches 1 - 10)

| Batch | Scope & Focus | Milestone Deliverables | Status |
| :---: | :--- | :--- | :---: |
| **Batch 1** | **Audit, Feature Inventory, Command Registry & Tool Catalog** | 12-Module Taxonomy, FeatureStatus, Roadmap Catalog, SuperSearch Launcher Upgrade, Registry Tests | **COMPLETED** |
| **Batch 2** | **Core Text Tools Expansion** | Reverse Text, Combine Cells, Clean Characters, Convert Numbers/Dates, Remove by Position, Spell Number | **COMPLETED** |
| **Batch 3** | **Range & Cell Operations** | Transpose, Flip H/V, Swap Ranges, Non-Contiguous Copy, Paste Visible Only, Merge Keep Data, Sign/Round | **COMPLETED** |
| **Batch 4** | **Find, Select, Duplicate Detection & Comparison** | Advanced Duplicate Finder, Select Duplicates, Multi-Sheet Find/Replace, Max/Min, Row Diff, Comments | **COMPLETED** |
| **Batch 5** | **Worksheet & Workbook Management** | Combine Sheets, Combine Workbooks, Split Workbook to Files, Sort Tabs, Sheet Index TOC, Delete Blanks | **COMPLETED** |
| **Batch 6** | **Filtering, Data Validation & Table Merge** | Advanced Multi-Sort, Color Filter/Count, Table Merge/Upsert, Dropdown Lists, Email/Phone Validation | **Next Up** |
| **Batch 7** | **Formatting, Formulas & Printing** | Reference Conversion (Abs/Rel), Exact Copy, Locate Errors, Math Operations, Number/Date Formats, Page Setup | Planned |
| **Batch 8** | **Import, Export & Picture Utilities** | CSV Import/Export, PDF Export, Batch Format Convert, Folder Listing, Cell Pictures, Chart Image Export | Planned |
| **Batch 9** | **Protection, Recovery, Charts & Analysis** | Batch Protect/Unprotect, Workbook Snapshots, Gantt/Waterfall Charts, Quick Statistics Dashboard | Planned |
| **Batch 10** | **Navigation, UX Polish, Packaging & Release** | Navigation Pane, Grid-Focus Crosshair, Date Picker, Formula Editor, Release Installer Verification | Planned |

---

## 7. Architecture & Compatibility Safeguards

1. **Excel 2016 Baseline Compatibility:**  
   Strict adherence to Excel 2016 COM API surfaces. Zero reliance on Office 365 `#SPILL!` dynamic array formulas or web-based Office.js APIs.
2. **Bulk Range I/O Performance (`Value2`):**  
   Reads and writes multi-cell ranges in single COM automation transfers using managed `object[,]` matrices. Managed transformation takes <500ms on 100K cells.
3. **Automated Excel State Restoration (`ExcelStateScope`):**  
   Suppresses `ScreenUpdating`, `DisplayAlerts`, `EnableEvents`, and switches `Calculation` to Manual during heavy operations. Each property is independently restored inside separate `try/catch` blocks in `Dispose()`.
4. **Data Integrity & Non-Destructive Defaults:**  
   Destination range collision checks prompt users before overwriting non-empty cells. High-impact operations (e.g. formula freezing) require explicit user confirmation.
5. **WPF Window Interop & High-DPI:**  
   All WPF dialogs are hosted as modal dialogs parented to Excel's native HWND via `WindowInteropHelper`, ensuring dialogs never fall behind the Excel host window across multi-monitor high-DPI scaling setups.
6. **Side-by-Side Coexistence:**  
   Completely independent COM Addin identity (`ExcelProductivitySuite`). Zero registry collision, file path overlap, or interference with existing Kutools installations.
