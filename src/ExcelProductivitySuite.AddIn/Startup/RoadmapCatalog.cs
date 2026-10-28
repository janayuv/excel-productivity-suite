using System;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.AddIn.Startup
{
    /// <summary>
    /// Central catalog registering comprehensive planned tools across all 12 modules (A through L)
    /// for instant discovery in the Tool Launcher and Feature Roadmap.
    /// </summary>
    public static class RoadmapCatalog
    {
        public static void RegisterPlannedTools(ICommandRegistry registry)
        {
            if (registry == null) throw new ArgumentNullException(nameof(registry));

            // =========================================================================
            // Module A: Text Tools (Remaining Future Roadmap)
            // =========================================================================
            AddPlanned(registry, "CMD_TEXT_BATCH_REPLACE", "Batch Replace Strings",
                "Execute multiple find-and-replace string rules simultaneously across the active selection.",
                CommandCategory.TextTools, "A", "Batch 2+",
                new[] { "batch replace", "multi replace", "substitute strings", "bulk replace", "regex replace" }, "ReplaceDialog");

            AddPlanned(registry, "CMD_TEXT_EXTRACT_NUMBERS", "Extract Numbers Only",
                "Extract only numeric digits and decimal points from mixed alphanumeric cell text.",
                CommandCategory.TextTools, "A", "Batch 2+",
                new[] { "extract numbers", "strip letters", "digits only", "isolate numbers", "extract integers" }, "Cut");

            AddPlanned(registry, "CMD_TEXT_EXTRACT_EMAILS", "Extract Email Addresses",
                "Locate and extract email addresses from free-form cell text into adjacent cells.",
                CommandCategory.TextTools, "A", "Batch 2+",
                new[] { "extract email", "find email", "isolate emails", "parse email address" }, "MailMergeSendEmail");

            AddPlanned(registry, "CMD_TEXT_EXTRACT_URLS", "Extract URLs and Links",
                "Extract web URLs (http/https) and domain links from selected cell text.",
                CommandCategory.TextTools, "A", "Batch 2+",
                new[] { "extract url", "find links", "parse websites", "http extract", "hyperlink text" }, "HyperlinkInsert");

            // =========================================================================
            // Module B: Range and Cell Operations (Remaining Future Roadmap)
            // =========================================================================
            AddPlanned(registry, "CMD_RANGE_COPY_NONCONTIGUOUS", "Copy Non-Contiguous Selection",
                "Copy multiple disjoint selection areas into a single contiguous destination range.",
                CommandCategory.RangeAndCells, "B", "Batch 3+",
                new[] { "copy noncontiguous", "disjoint copy", "multi area copy", "ctrl select copy" }, "Copy");

            AddPlanned(registry, "CMD_RANGE_PASTE_VISIBLE_ONLY", "Paste into Visible Cells Only",
                "Paste clipboard data strictly into visible cells, bypassing filtered or hidden rows.",
                CommandCategory.RangeAndCells, "B", "Batch 3+",
                new[] { "paste visible", "skip hidden rows", "paste filtered", "visible cells only" }, "Paste");

            AddPlanned(registry, "CMD_RANGE_UNMERGE_FILL", "Unmerge and Fill Cells",
                "Safely unmerge composite cells and automatically populate all constituent cells with the original value.",
                CommandCategory.RangeAndCells, "B", "Batch 3+",
                new[] { "unmerge and fill", "split merged", "fill unmerged", "normalize merged cells" }, "TableSplitCells");

            AddPlanned(registry, "CMD_RANGE_RANDOM_SAMPLE", "Random Sampling of Records",
                "Select or extract a random sample of N rows or P% of records from a dataset.",
                CommandCategory.RangeAndCells, "B", "Batch 3+",
                new[] { "random sample", "sample rows", "extract sample", "audit sample", "statistical sampling" }, "FunctionsStatisticalInsertGallery");

            // =========================================================================
            // Module C: Find, Select, and Compare (Remaining Future Roadmap)
            // =========================================================================
            AddPlanned(registry, "CMD_FIND_BY_FORMAT", "Find Cells by Format & Color",
                "Locate cells matching specific font color, fill color, font style, or border attributes.",
                CommandCategory.FindSelectCompare, "C", "Batch 4+",
                new[] { "find by color", "find formatting", "match cell color", "find bold cells" }, "ColorPicker");

            AddPlanned(registry, "CMD_FIND_BY_VALIDATION", "Find Cells with Data Validation",
                "Select and highlight all cells that have Excel data validation rules applied.",
                CommandCategory.FindSelectCompare, "C", "Batch 4+",
                new[] { "find validation", "cells with dropdowns", "validation cells", "inspect validation" }, "DataValidation");

            // =========================================================================
            // Module D: Worksheet and Workbook Management (Batch 5)
            // =========================================================================
            AddPlanned(registry, "CMD_SHEET_COMBINE", "Combine Worksheets into One",
                "Merge multiple worksheets into a single consolidated master sheet with header options.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "combine sheets", "merge worksheets", "stack sheets", "master worksheet", "consolidate sheets" }, "TableMergeCells");

            AddPlanned(registry, "CMD_WB_COMBINE_FILES", "Combine Workbooks",
                "Combine sheets from multiple external Excel workbooks into one central workbook.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "combine workbooks", "merge excel files", "batch combine", "consolidate files" }, "FileOpen");

            AddPlanned(registry, "CMD_WB_SPLIT_FILES", "Split Workbook into Separate Files",
                "Split each worksheet of the active workbook into an independent Excel file (.xlsx, .pdf, .csv).",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "split workbook", "save sheets as files", "separate workbooks", "export sheets" }, "TableSplitCells");

            AddPlanned(registry, "CMD_SHEET_SORT_TABS", "Sort Worksheet Tabs",
                "Sort sheet tabs alphabetically (A-Z or Z-A), numerically, or by tab color.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "sort sheet tabs", "sort worksheets", "alphabetical sheets", "tab order", "reorder sheets" }, "SortAscending");

            AddPlanned(registry, "CMD_SHEET_CREATE_INDEX", "Create Linked Sheet Index (TOC)",
                "Generate a table of contents worksheet with clickable hyperlinks to every sheet in the workbook.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "sheet index", "table of contents", "linked sheets", "toc generator", "worksheet directory" }, "HyperlinkInsert");

            AddPlanned(registry, "CMD_SHEET_DELETE_BLANKS", "Delete Blank Worksheets",
                "Identify and delete all completely empty worksheets in the active workbook with confirmation.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "delete blank sheets", "remove empty worksheets", "clean empty sheets" }, "SheetDelete");

            AddPlanned(registry, "CMD_SHEET_COPY_PAGE_SETUP", "Copy Page Setup Across Sheets",
                "Copy margins, orientation, headers/footers, and print scaling from active sheet to other sheets.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "copy page setup", "sync margins", "duplicate print settings", "header footer copy" }, "PageSetupPageDialog");

            AddPlanned(registry, "CMD_SHEET_FREEZE_MULTI", "Freeze / Unfreeze Panes Across Sheets",
                "Apply or remove Freeze Panes simultaneously across all or selected worksheets in the workbook.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "freeze multiple sheets", "unfreeze all", "sync freeze panes", "batch freeze" }, "WindowFreezePanes");

            AddPlanned(registry, "CMD_SHEET_SEQUENCE_NAMES", "Create Sequentially Named Sheets",
                "Generate a series of new worksheets with sequential names (e.g. Month_01 to Month_12).",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "sequential sheets", "create multiple sheets", "numbered sheets", "batch add sheets" }, "SheetInsert");

            AddPlanned(registry, "CMD_SHEET_CONSOLIDATE", "Consolidate Data Across Sheets",
                "Aggregate tabular data across identical structured sheets using SUM, AVERAGE, or COUNT.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "consolidate sheets", "3d consolidation", "sum across sheets", "multi sheet rollup" }, "Consolidate");

            AddPlanned(registry, "CMD_WB_OPEN_FOLDER", "Open Containing Folder",
                "Open the Windows File Explorer folder containing the active workbook.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "open folder", "containing folder", "explore directory", "reveal in explorer" }, "FileOpen");

            AddPlanned(registry, "CMD_WB_COPY_PATH", "Copy Workbook Full Path",
                "Copy the full file path of the active workbook to the Windows clipboard.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "copy path", "copy file location", "workbook address", "clipboard path" }, "Copy");

            AddPlanned(registry, "CMD_WB_PROPERTIES", "Workbook Extended Properties",
                "View file size, creation timestamp, author, sheets count, and COM metadata.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "workbook properties", "file info", "stats", "sheet count", "document metadata" }, "Info");

            // =========================================================================
            // Module E: Data, Filtering, and Statistics (Batch 6)
            // =========================================================================
            AddPlanned(registry, "CMD_DATA_ADVANCED_SORT", "Advanced Multi-Level Sort",
                "Sort table by multiple custom criteria including text length, month/weekday name, and colors.",
                CommandCategory.DataAndFiltering, "E", "Batch 6",
                new[] { "advanced sort", "multi sort", "sort by length", "custom sort order" }, "SortDialog");

            AddPlanned(registry, "CMD_DATA_FILTER_BY_COLOR", "Filter by Cell / Font Color",
                "Filter tabular data based on specific background fill colors or font colors.",
                CommandCategory.DataAndFiltering, "E", "Batch 6",
                new[] { "filter by color", "color filter", "fill color filter", "highlight filter" }, "Filter");

            AddPlanned(registry, "CMD_DATA_COUNT_BY_COLOR", "Count and Sum by Color",
                "Calculate total counts, sums, and averages for cells sharing identical background or font colors.",
                CommandCategory.DataAndFiltering, "E", "Batch 6",
                new[] { "count by color", "sum by color", "color statistics", "color aggregation" }, "FunctionsStatisticalInsertGallery");

            AddPlanned(registry, "CMD_DATA_MERGE_TABLES", "Merge / Update Tables on Keys",
                "Upsert and update a target table using matched key columns from a secondary lookup table.",
                CommandCategory.DataAndFiltering, "E", "Batch 6",
                new[] { "merge tables", "vlookup merge", "table upsert", "join tables", "update records" }, "TableMergeCells");

            AddPlanned(registry, "CMD_DATA_QUICK_SUBTOTAL", "Generate Summaries & Subtotals",
                "Insert grouping subtotals and summaries at every change in designated categorical columns.",
                CommandCategory.DataAndFiltering, "E", "Batch 6",
                new[] { "subtotals", "group summaries", "data subtotals", "category rollup" }, "PivotTableSubtotals");

            AddPlanned(registry, "CMD_DATA_FREQUENCY_STATS", "Value Frequency Distribution",
                "Generate a statistical frequency table showing distinct values, occurrences, and percentages.",
                CommandCategory.DataAndFiltering, "E", "Batch 6",
                new[] { "frequency distribution", "value count", "histogram table", "occurrence stats" }, "FunctionsStatisticalInsertGallery");

            AddPlanned(registry, "CMD_DATA_EXTRACT_DISTINCT", "Extract Unique Records",
                "Extract distinct/unique records from a dataset and write to a clean destination table.",
                CommandCategory.DataAndFiltering, "E", "Batch 6",
                new[] { "extract unique", "distinct records", "filter unique", "isolate distinct" }, "TableDeleteRows");

            // =========================================================================
            // Module F: Data Validation and Dropdown Lists (Batch 6)
            // =========================================================================
            AddPlanned(registry, "CMD_VALIDATION_CREATE_DROPDOWN", "Create Standard Dropdown List",
                "Create clean Excel in-cell dropdown validation lists from selected source items or ranges.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "create dropdown", "data validation list", "cell dropdown", "combo list" }, "DataValidation");

            AddPlanned(registry, "CMD_VALIDATION_SEARCHABLE_DROPDOWN", "Searchable Dropdown Interface",
                "Provide an enhanced searchable popup list for fast data entry in cells with large validation lists.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "searchable dropdown", "autocomplete dropdown", "fast validation picker", "filter list" }, "FindDialog");

            AddPlanned(registry, "CMD_VALIDATION_MULTI_SELECT", "Multi-Selection Dropdown List",
                "Allow selecting multiple items from a dropdown list separated by commas within a single cell.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "multi select dropdown", "multiple choices", "comma separated dropdown", "multi pick" }, "CheckMark");

            AddPlanned(registry, "CMD_VALIDATION_PREVENT_DUPS", "Prevent Duplicate Entries Rule",
                "Apply custom validation rules that reject duplicate entries in real-time during user input.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "prevent duplicates", "unique validation", "reject dups", "restrict duplicates" }, "ShieldAlert");

            AddPlanned(registry, "CMD_VALIDATION_EMAIL_RULE", "Validate Email Format",
                "Apply regex-based validation rules ensuring user input adheres to valid email address syntax.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "email validation", "validate email", "email format rule", "restrict email" }, "MailMergeSendEmail");

            AddPlanned(registry, "CMD_VALIDATION_PHONE_RULE", "Validate Phone Format",
                "Apply input validation rules enforcing telephone number patterns.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "phone validation", "validate phone", "telephone pattern", "restrict phone" }, "Call");

            AddPlanned(registry, "CMD_VALIDATION_DATE_RANGE", "Restrict Date Range",
                "Enforce minimum and maximum allowable dates with custom error warning messages.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "restrict date", "date validation", "valid dates only", "date boundary" }, "FormatDate");

            AddPlanned(registry, "CMD_VALIDATION_HIGHLIGHT_INVALID", "Highlight Invalid Validation Entries",
                "Scan range and highlight all cells whose current values violate existing data validation rules.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "highlight invalid", "validation errors", "invalid cells", "circle invalid data" }, "ShieldAlert");

            AddPlanned(registry, "CMD_VALIDATION_BULK_REMOVE", "Remove Validation Rules in Bulk",
                "Clear all data validation constraints across selected ranges or entire worksheets.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "clear validation", "remove dropdowns", "delete validation rules", "reset validation" }, "Delete");

            AddPlanned(registry, "CMD_VALIDATION_COPY_RULES", "Copy Validation Rules",
                "Copy data validation specifications from a source range to multiple target ranges.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "copy validation", "duplicate validation rules", "sync dropdowns" }, "Copy");

            // =========================================================================
            // Module G: Formula and Calculation Utilities (Batch 7)
            // =========================================================================
            AddPlanned(registry, "CMD_FORMULA_CONVERT_REF", "Convert References (Abs/Rel)",
                "Toggle cell references between Absolute ($A$1), Row Absolute (A$1), Col Absolute ($A1), and Relative (A1).",
                CommandCategory.FormulasAndCalculation, "G", "Batch 7",
                new[] { "convert references", "absolute relative", "f4 convert", "toggle reference type", "fix references" }, "FormulaMoreFunctionsMenu");

            AddPlanned(registry, "CMD_FORMULA_EXACT_COPY", "Exact Formula Copy",
                "Copy formulas to new locations exactly without shifting relative cell references.",
                CommandCategory.FormulasAndCalculation, "G", "Batch 7",
                new[] { "exact copy", "copy formula without changing", "freeze references copy", "duplicate formula" }, "Copy");

            AddPlanned(registry, "CMD_FORMULA_AUDIT_PRECEDENTS", "Audit Precedents Visually",
                "Navigate and inspect all precedent cells contributing to the active cell formula.",
                CommandCategory.FormulasAndCalculation, "G", "Batch 7",
                new[] { "trace precedents", "audit formula", "precedent cells", "formula dependency" }, "TracePrecedents");

            AddPlanned(registry, "CMD_FORMULA_AUDIT_DEPENDENTS", "Audit Dependents Visually",
                "Navigate and inspect all cells dependent upon the active cell value.",
                CommandCategory.FormulasAndCalculation, "G", "Batch 7",
                new[] { "trace dependents", "dependent cells", "formula dependents", "what uses this cell" }, "TraceDependents");

            AddPlanned(registry, "CMD_FORMULA_LOCATE_ERRORS", "Locate Formula Errors",
                "Find, select, and navigate through all formula error cells (#N/A, #VALUE!, #REF!, #DIV/0!).",
                CommandCategory.FormulasAndCalculation, "G", "Batch 7",
                new[] { "locate errors", "find formula errors", "navigate errors", "fix #N/A", "error finder" }, "ErrorChecking");

            AddPlanned(registry, "CMD_FORMULA_BATCH_REPLACE_REF", "Batch Replace Formula References",
                "Substitute sheet names or range addresses within formulas across all selected cells.",
                CommandCategory.FormulasAndCalculation, "G", "Batch 7",
                new[] { "replace references", "change formula sheet", "update formula ranges", "batch formula edit" }, "ReplaceDialog");

            AddPlanned(registry, "CMD_FORMULA_MATH_OPERATIONS", "Apply Math Operation to Cells",
                "Add, subtract, multiply, or divide a fixed constant across all selected numeric cells.",
                CommandCategory.FormulasAndCalculation, "G", "Batch 7",
                new[] { "math operation", "multiply range", "divide range", "add constant", "bulk calculation" }, "Calculator");

            AddPlanned(registry, "CMD_FORMULA_EXPLAIN", "Explain Formula Deterministically",
                "Deconstruct and explain nested formulas in plain English using a local deterministic parser.",
                CommandCategory.FormulasAndCalculation, "G", "Batch 7",
                new[] { "explain formula", "formula parser", "deconstruct formula", "understand formula" }, "Info");

            AddPlanned(registry, "CMD_FORMULA_TEMPLATES", "Common Formula Templates",
                "Insert pre-built formulas for CAGR, working days, loan amortization, and standard metrics.",
                CommandCategory.FormulasAndCalculation, "G", "Batch 7",
                new[] { "formula templates", "financial formulas", "date formulas", "insert formula" }, "FunctionsFinancialInsertGallery");

            // =========================================================================
            // Module H: Formatting and Design (Batch 7)
            // =========================================================================
            AddPlanned(registry, "CMD_FORMAT_BATCH_NUMBERS", "Batch Number & Currency Formatting",
                "Apply standardized financial, decimal, percentage, and currency masks across ranges.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "number format", "currency format", "decimal places", "standardize formats", "accounting format" }, "FormatNumberCurrency");

            AddPlanned(registry, "CMD_FORMAT_DATE_STANDARDS", "Standardize Date Formats",
                "Format date cells to ISO 8601 (YYYY-MM-DD), Month Day Year, or local enterprise styles.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "date format", "iso date", "standardize dates", "format date cells" }, "FormatDate");

            AddPlanned(registry, "CMD_FORMAT_HIGHLIGHT_CELL_TYPES", "Highlight Cell Types",
                "Apply distinct background colors to formulas, constants, locked cells, errors, and text.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "highlight cell types", "color formulas", "color constants", "visual cell audit" }, "ColorPicker");

            AddPlanned(registry, "CMD_FORMAT_BORDERS", "Configurable Border Schemes",
                "Apply elegant table borders, gridlines, thick headers, or double bottom totals.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "table borders", "border styles", "accounting borders", "gridlines" }, "BorderAll");

            AddPlanned(registry, "CMD_FORMAT_ROW_HEIGHT", "Uniform Row Heights",
                "Set uniform row heights or smart auto-fit across multiple tables with minimum thresholds.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "row height", "uniform rows", "auto fit row", "standardize height" }, "TableRowsInsertAbove");

            AddPlanned(registry, "CMD_FORMAT_COL_WIDTH", "Uniform Column Widths",
                "Set standardized column widths or auto-fit with extra padding across selections.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "column width", "uniform columns", "auto fit columns", "standardize width" }, "TableColumnsInsertLeft");

            AddPlanned(registry, "CMD_PRINT_SET_PRINT_AREAS", "Set Multiple Print Areas",
                "Configure and manage print areas across multiple selected ranges or sheets simultaneously.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "print area", "set print areas", "page layout", "print setup" }, "PrintAreaMenu");

            AddPlanned(registry, "CMD_PRINT_INSERT_PAGE_BREAKS", "Insert Page Breaks by Rule",
                "Automatically insert horizontal page breaks every N rows or when key column values change.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "insert page breaks", "page breaks every n rows", "paging by category", "print breaks" }, "PageBreakInsert");

            AddPlanned(registry, "CMD_PRINT_REPEATING_HEADERS", "Configure Repeating Print Headers",
                "Set title rows and columns to repeat on every printed page across multiple sheets.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "repeating headers", "print titles", "rows to repeat at top", "print header rows" }, "PrintTitles");

            AddPlanned(registry, "CMD_PRINT_COPY_PAGE_SETUP", "Batch Copy Page Settings",
                "Synchronize orientation, paper size, margins, and scaling across all workbook sheets.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "copy page settings", "sync page setup", "batch print configuration" }, "PageSetupPageDialog");

            // =========================================================================
            // Module I: Import, Export, and File Utilities (Batch 8)
            // =========================================================================
            AddPlanned(registry, "CMD_FILE_IMPORT_CSV", "Import CSV / Text Files",
                "Import delimited text and CSV files with custom encodings, delimiters, and column datatypes.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8",
                new[] { "import csv", "import text", "load delimited file", "read csv into sheet" }, "GetExternalDataFromText");

            AddPlanned(registry, "CMD_FILE_EXPORT_RANGE_CSV", "Export Range to CSV",
                "Export selected cells or active table to a standalone CSV, TSV, or JSON file.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8",
                new[] { "export csv", "save range as csv", "export selection", "export tsv" }, "SaveAsCsv");

            AddPlanned(registry, "CMD_FILE_EXPORT_SHEETS_PDF", "Export Worksheets to PDF",
                "Batch export selected worksheets or entire workbook to high-resolution PDF documents.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8",
                new[] { "export pdf", "save as pdf", "batch pdf export", "convert sheets to pdf" }, "FileSaveAsPdfOrXps");

            AddPlanned(registry, "CMD_FILE_BATCH_CONVERT", "Batch Convert Workbook Formats",
                "Convert folders of legacy .xls files to modern .xlsx or export workbooks to CSV.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8",
                new[] { "batch convert workbooks", "convert xls to xlsx", "bulk format conversion" }, "FileSaveAs");

            AddPlanned(registry, "CMD_FILE_FOLDER_LISTING", "Generate File & Folder Inventory",
                "Scan a directory on disk and output a comprehensive file inventory table with file sizes and dates.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8",
                new[] { "folder listing", "file inventory", "list files in folder", "directory scan" }, "FileOpen");

            AddPlanned(registry, "CMD_FILE_IMPORT_PICTURES", "Import Pictures into Cells",
                "Batch import image files into cells with automated sizing and cell locking.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8",
                new[] { "import pictures", "insert images in cells", "batch photos", "fit image to cell" }, "PictureInsertFromFile");

            AddPlanned(registry, "CMD_FILE_MATCH_PICTURES", "Match Pictures by Key",
                "Match image file names against product codes or employee IDs in cells and insert pictures.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8",
                new[] { "match pictures", "lookup images", "insert pictures by name", "product photos" }, "PictureInsertFromFile");

            AddPlanned(registry, "CMD_FILE_EXPORT_CHARTS", "Export Charts to Images",
                "Export all embedded charts and shapes in the workbook as PNG, JPEG, or SVG files.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8",
                new[] { "export charts", "save chart as image", "chart picture export", "png charts" }, "ChartTypePieInsertGallery");

            AddPlanned(registry, "CMD_FILE_INSERT_TEXT_FILE", "Insert Text File Content",
                "Insert raw content from an external text file at the current active cell cursor.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8",
                new[] { "insert text file", "append file text", "read text into cell" }, "FileOpen");

            // =========================================================================
            // Module J: Protection and Recovery (Batch 9)
            // =========================================================================
            AddPlanned(registry, "CMD_SECURITY_BATCH_PROTECT", "Batch Protect Worksheets",
                "Apply protection passwords and permission flags across multiple worksheets simultaneously.",
                CommandCategory.ProtectionAndRecovery, "J", "Batch 9",
                new[] { "batch protect", "protect multiple sheets", "lock worksheets", "sheet security" }, "ProtectSheet");

            AddPlanned(registry, "CMD_SECURITY_BATCH_UNPROTECT", "Batch Unprotect Worksheets",
                "Batch unprotect multiple worksheets using a single user-authorized password entry.",
                CommandCategory.ProtectionAndRecovery, "J", "Batch 9",
                new[] { "batch unprotect", "unlock multiple sheets", "unprotect sheets", "remove password" }, "UnprotectSheet");

            AddPlanned(registry, "CMD_SECURITY_WORKBOOK_ENCRYPT", "Manage Workbook Password",
                "Securely set or clear workbook open and modify passwords.",
                CommandCategory.ProtectionAndRecovery, "J", "Batch 9",
                new[] { "workbook password", "encrypt workbook", "protect workbook structure" }, "ProtectWorkbook");

            AddPlanned(registry, "CMD_SECURITY_WINDOW_VISIBILITY", "Manage Workbook Windows",
                "Batch hide or unhide workbook windows across multi-document sessions.",
                CommandCategory.ProtectionAndRecovery, "J", "Batch 9",
                new[] { "hide windows", "unhide window", "window visibility", "excel windows" }, "WindowHide");

            AddPlanned(registry, "CMD_RECOVERY_CREATE_SNAPSHOT", "Create Workbook Snapshot",
                "Take a temporary in-session snapshot of active workbook state for safe trial operations.",
                CommandCategory.ProtectionAndRecovery, "J", "Batch 9",
                new[] { "workbook snapshot", "take snapshot", "backup state", "undo checkpoint", "restore point" }, "FileSaveAs");

            AddPlanned(registry, "CMD_RECOVERY_RESTORE_SNAPSHOT", "Restore Workbook Snapshot",
                "Revert the active workbook back to a previously saved in-session snapshot safely.",
                CommandCategory.ProtectionAndRecovery, "J", "Batch 9",
                new[] { "restore snapshot", "revert workbook", "undo to checkpoint", "restore state" }, "FileOpen");

            // =========================================================================
            // Module K: Charts and Analysis (Batch 9)
            // =========================================================================
            AddPlanned(registry, "CMD_CHART_TEMPLATES", "Reusable Chart Templates",
                "Apply standardized corporate color palettes and clean formatting to selected charts.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9",
                new[] { "chart templates", "format chart", "corporate charts", "clean chart style" }, "ChartTypeColumnInsertGallery");

            AddPlanned(registry, "CMD_CHART_GANTT", "Generate Project Gantt Chart",
                "Create a project timeline Gantt chart from task names, start dates, and durations.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9",
                new[] { "gantt chart", "project timeline", "schedule chart", "task bar chart" }, "ChartTypeBarInsertGallery");

            AddPlanned(registry, "CMD_CHART_WATERFALL", "Generate Waterfall Chart",
                "Construct a bridge / waterfall variance chart compatible with Excel 2016.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9",
                new[] { "waterfall chart", "bridge chart", "variance analysis", "financial chart" }, "ChartTypeColumnInsertGallery");

            AddPlanned(registry, "CMD_CHART_TARGET_VS_ACTUAL", "Target vs. Actual Chart",
                "Generate bullet or target-versus-actual comparison charts with achievement thresholds.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9",
                new[] { "target vs actual", "bullet chart", "kpi chart", "performance chart" }, "ChartTypeLineInsertGallery");

            AddPlanned(registry, "CMD_ANALYSIS_QUICK_STATS", "Quick Statistics Dashboard",
                "Display an instant audit summary with Count, Sum, Mean, Median, Min, Max, and StdDev.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9",
                new[] { "quick stats", "summary statistics", "median mean", "descriptive statistics" }, "FunctionsStatisticalInsertGallery");

            AddPlanned(registry, "CMD_ANALYSIS_PIVOT_HELPERS", "PivotTable Utilities",
                "Bulk refresh all PivotTables, adjust field formatting, and toggle classic layout.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9",
                new[] { "pivot table helpers", "refresh pivot tables", "pivot formatting", "classic pivot" }, "PivotTableInsert");

            // =========================================================================
            // Module L: Navigation and User Productivity (Batch 10)
            // =========================================================================
            AddPlanned(registry, "CMD_NAV_PANE", "Workbook & Sheet Navigation Pane",
                "Docked task pane providing rapid navigation across workbooks, sheets, and tables.",
                CommandCategory.NavigationAndProductivity, "L", "Batch 10",
                new[] { "navigation pane", "sheet sidebar", "fast navigation", "sheet list panel" }, "ViewNormalViewExcel");

            AddPlanned(registry, "CMD_NAV_NAMED_RANGES", "Named Range Navigator",
                "Search, inspect, navigate, and edit named ranges across the entire workbook.",
                CommandCategory.NavigationAndProductivity, "L", "Batch 10",
                new[] { "named ranges", "name manager", "locate named range", "inspect names" }, "NameManager");

            AddPlanned(registry, "CMD_NAV_COLUMNS", "Column Navigator & Visibility",
                "Quickly locate, hide, unhide, and navigate across wide spreadsheets with 50+ columns.",
                CommandCategory.NavigationAndProductivity, "L", "Batch 10",
                new[] { "column navigator", "find column", "wide sheet navigation", "hide columns list" }, "TableColumnsInsertLeft");

            AddPlanned(registry, "CMD_NAV_CELL_PREVIEW", "Expanded Cell Content Viewer",
                "View and scroll long multiline text or formulas in a dedicated modal inspection card.",
                CommandCategory.NavigationAndProductivity, "L", "Batch 10",
                new[] { "cell preview", "view long cell", "multiline viewer", "inspect text" }, "ReviewShowOrHideMarkup");

            AddPlanned(registry, "CMD_NAV_FORMULA_EDITOR", "Expanded Multi-Line Formula Editor",
                "Edit long, complex nested formulas with syntax highlighting, indentation, and formatting.",
                CommandCategory.NavigationAndProductivity, "L", "Batch 10",
                new[] { "expanded formula editor", "multiline formula", "big formula editor", "format formula" }, "FormulaMoreFunctionsMenu");

            AddPlanned(registry, "CMD_NAV_DATE_PICKER", "Interactive Date Picker",
                "Insert dates into cells using a floating calendar date picker interface.",
                CommandCategory.NavigationAndProductivity, "L", "Batch 10",
                new[] { "date picker", "calendar picker", "insert date popup", "cell calendar" }, "FormatDate");

            AddPlanned(registry, "CMD_NAV_GRID_FOCUS", "Grid-Focus Reading Crosshair",
                "Highlight the active row and column dynamically to prevent row reading errors in wide tables.",
                CommandCategory.NavigationAndProductivity, "L", "Batch 10",
                new[] { "grid focus", "crosshair highlighting", "reading mode", "highlight active row" }, "SelectCell");

            AddPlanned(registry, "CMD_NAV_DATA_ENTRY_FORM", "Structured Record Entry Form",
                "Display a dynamic modal form for structured row-by-row data entry with validation.",
                CommandCategory.NavigationAndProductivity, "L", "Batch 10",
                new[] { "data entry form", "record form", "add row dialog", "structured input" }, "FormMenu");
        }

        private static void AddPlanned(
            ICommandRegistry registry,
            string id,
            string title,
            string description,
            CommandCategory category,
            string moduleCode,
            string roadmapBatch,
            string[] keywords,
            string? officeIconName = null)
        {
            registry.RegisterCommand(new SuiteCommand(
                id: id,
                title: title,
                description: description,
                category: category,
                executeAction: ctx =>
                {
                    ctx.ShowMessage(
                        $"The '{title}' tool is part of Module {moduleCode} ({category}) and is scheduled for implementation in {roadmapBatch}.\n\nFull specifications and requirements are tracked in the Feature Inventory.",
                        $"{title} - Roadmap",
                        isError: false);
                },
                keywords: keywords,
                officeIconName: officeIconName,
                status: FeatureStatus.Planned,
                moduleCode: moduleCode,
                roadmapBatch: roadmapBatch
            ));
        }
    }
}
