using System;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.UI.Dialogs;

namespace ExcelProductivitySuite.AddIn.Startup
{
    public static class ToolRegistrar
    {
        public static void RegisterAllTools(ICommandRegistry registry)
        {
            if (registry == null) throw new ArgumentNullException(nameof(registry));

            // 1. Tool Launcher (SuperSearch) - Module L / General
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_LAUNCHER",
                title: "Tool Launcher",
                description: "Search, filter, and launch any productivity utility instantly.",
                category: CommandCategory.NavigationAndProductivity,
                executeAction: ctx =>
                {
                    var dlg = new ToolLauncherWindow(registry, ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "search", "launcher", "find tool", "palette", "commands" },
                shortcutHint: "Ctrl+Shift+P",
                officeIconName: "FindDialog",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "L",
                roadmapBatch: "Batch 1"
            ));

            // 2. Extract Text - Module A: Text Tools
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_EXTRACT",
                title: "Extract Text",
                description: "Extract substrings before, after, or between delimiters into cells.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new ExtractTextDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "extract", "substring", "delimiter", "parse", "slice" },
                officeIconName: "Cut",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Active"
            ));

            // 3. Split Text - Module A: Text Tools
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_SPLIT",
                title: "Split Cells",
                description: "Split delimited cell values across multiple columns or into new rows.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new SplitTextDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "split", "columns", "rows", "delimiter", "tokenize" },
                officeIconName: "TableSplitCells",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Active"
            ));

            // 4. Remove Characters - Module A: Text Tools
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_REMOVE_CHARS",
                title: "Remove Characters",
                description: "Strip numeric, alphabetic, symbols, or custom character sets.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new RemoveCharactersDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "remove", "delete characters", "strip digits", "clean symbols" },
                officeIconName: "Delete",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Active"
            ));

            // 5. Trim & Normalize Whitespace - Module A: Text Tools
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_TRIM_SPACES",
                title: "Trim Spaces",
                description: "Clean excess spaces, leading/trailing gaps, and non-breaking spaces.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new WhitespaceDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "trim", "whitespace", "spaces", "clean", "nbsp" },
                officeIconName: "ParagraphSpacingAfter",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Active"
            ));

            // 6. Change Text Case - Module A: Text Tools
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_CHANGE_CASE",
                title: "Change Case",
                description: "Convert text to UPPERCASE, lowercase, Proper Case, or Sentence case.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new ChangeCaseDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "case", "uppercase", "lowercase", "proper case", "capitalize" },
                officeIconName: "ChangeCase",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Active"
            ));

            // 7. Add Text / Affixes - Module A: Text Tools
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_ADD_AFFIX",
                title: "Add Text",
                description: "Prepend prefix, append suffix, or insert text at position with safeguards.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new AddTextDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "prefix", "suffix", "append", "prepend", "affix", "insert position" },
                officeIconName: "AdpAddExisting",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Active"
            ));

            // 8. Combine Cells - Module A: Text Tools (Batch 2)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_COMBINE",
                title: "Combine Cells",
                description: "Combine text across columns, down rows, or entire selection with custom delimiters.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new CombineCellsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "combine", "merge text", "concatenate", "join cells", "rows", "columns" },
                officeIconName: "GroupRows",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Batch 2"
            ));

            // 9. Reverse Text - Module A: Text Tools (Batch 2)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_REVERSE",
                title: "Reverse Text",
                description: "Reverse text string order by characters (surrogate-safe), words, or custom delimiters.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new ReverseTextDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "reverse text", "flip text", "reverse words", "backwards", "invert string" },
                officeIconName: "SortUp",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Batch 2"
            ));

            // 10. Remove by Position - Module A: Text Tools (Batch 2)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_REMOVE_BY_POS",
                title: "Remove by Position",
                description: "Remove characters from left, right, middle range, or before/after delimiter.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new RemoveByPositionDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "remove by position", "truncate", "strip left", "strip right", "substring remove" },
                officeIconName: "Cut",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Batch 2"
            ));

            // 11. Clean Non-Printing Characters - Module A: Text Tools (Batch 2)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_CLEAN_NONPRINTING",
                title: "Clean Characters",
                description: "Remove ASCII control characters, zero-width spaces, and decode HTML entities.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new CleanTextDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "clean text", "control characters", "nonprinting", "zero width", "html decode", "sanitize" },
                officeIconName: "Clear",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Batch 2"
            ));

            // 12. Convert Numbers Stored as Text - Module A: Text Tools (Batch 2)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_CONVERT_NUMBERS",
                title: "Convert to Numbers",
                description: "Convert numbers stored as text, currency symbols, accounting parens, and percentages.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new ConvertNumbersDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "convert numbers", "text to number", "strip currency", "accounting format", "parse numeric" },
                officeIconName: "NumberFormatCurrency",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Batch 2"
            ));

            // 13. Convert Text Dates - Module A: Text Tools (Batch 2)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_CONVERT_DATES",
                title: "Convert Dates",
                description: "Parse and convert varied text date formats into standardized ISO, EU, US, or long dates.",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new ConvertDatesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "convert dates", "parse dates", "iso date", "text to date", "reformat dates" },
                officeIconName: "TimeDateInsert",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Batch 2"
            ));

            // 14. Spell Numbers - Module A: Text Tools (Batch 2)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_SPELL_NUMBER",
                title: "Spell Numbers",
                description: "Convert numbers and currencies into written English words (e.g. check amounts).",
                category: CommandCategory.TextTools,
                executeAction: ctx =>
                {
                    var dlg = new SpellNumberDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "spell number", "number to words", "cheque amount", "currency words", "cardinal numbers" },
                officeIconName: "Spelling",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Batch 2"
            ));

            // 8. Fill Blank Cells - Module B: Range and Cell Operations
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_FILL_BLANKS",
                title: "Fill Blank Cells",
                description: "Fill empty cells with values from above, below, left, right, or constant.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new FillBlanksDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "fill blanks", "fill down", "empty cells", "interpolate" },
                officeIconName: "FillDown",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Active"
            ));

            // Batch 3 - Module B: Range and Cell Operations Expansion
            // Transpose Range
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_TRANSPOSE",
                title: "Transpose Range",
                description: "Transpose rows to columns and columns to rows in-place, at target cell, or new sheet.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new TransposeRangeDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "transpose", "rows to columns", "columns to rows", "flip matrix", "pivot dimensions" },
                officeIconName: "PivotTableSubtotals",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            // Flip Range
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_FLIP_HORIZONTAL",
                title: "Flip Range",
                description: "Reverse column order horizontally left-to-right or reverse row order vertically top-to-bottom.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new FlipRangeDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "flip range", "flip horizontal", "flip vertical", "reverse columns", "reverse rows", "mirror" },
                officeIconName: "ViewSideBySide",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            // Swap Ranges
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_SWAP",
                title: "Swap Ranges",
                description: "Exchange the values and formatting between two selected non-overlapping ranges.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new SwapRangesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "swap ranges", "exchange cells", "switch ranges", "swap columns", "swap data" },
                officeIconName: "TableMergeCells",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            // Change Sign
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_CHANGE_SIGN",
                title: "Change Sign (+/-)",
                description: "Invert or enforce positive/negative signs on all numeric cells in the selected range.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new ChangeSignDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "change sign", "invert sign", "make negative", "make positive", "reverse sign", "absolute value" },
                officeIconName: "Calculator",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            // Round Values Without Formulas
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_ROUND_VALUES",
                title: "Round Values",
                description: "Round, ceil, floor, truncate, or snap numbers to multiples in-place without formulas.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new RoundValuesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "round values", "round decimals", "truncate decimals", "ceiling", "floor", "round in place" },
                officeIconName: "FormatNumberCurrency",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            // Randomly Sort Range
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_RANDOM_SORT",
                title: "Random Sort",
                description: "Shuffle rows, columns, or individual cells in randomized order using Fisher-Yates.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new RandomSortDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "random sort", "shuffle rows", "randomize order", "random list", "shuffle columns" },
                officeIconName: "SortAscending",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            // Merge Cells (Keep All Data)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_MERGE_KEEP_DATA",
                title: "Merge (Keep Data)",
                description: "Merge cells into one visual block while preserving all cell contents with an explicit separator.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new MergeKeepDataDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "merge keep data", "merge preserve text", "combine and merge", "non destructive merge" },
                officeIconName: "TableMergeCells",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            // Insert Blank Rows by Rule
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_INSERT_BLANK_ROWS",
                title: "Insert Blank Rows",
                description: "Insert blank rows at regular intervals (e.g. every N rows) or upon key column changes.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new InsertBlankRowsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "insert blank rows", "interval rows", "spacing rows", "add empty rows", "value change rows" },
                officeIconName: "TableRowsInsertAbove",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            // Delete Blank Rows / Columns
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_DELETE_BLANK_ROWS",
                title: "Delete Blank Rows",
                description: "Scan selection or sheet and batch delete completely empty rows or columns safely.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new DeleteBlankRowsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "delete blank rows", "remove empty rows", "delete empty columns", "clean blank rows" },
                officeIconName: "TableDeleteRows",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            // 9. Compare Ranges - Module C: Find, Select, and Compare
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_COMPARE",
                title: "Compare Ranges",
                description: "Identify cell differences, mismatches, or missing items between two ranges or sheets.",
                category: CommandCategory.FindSelectCompare,
                executeAction: ctx =>
                {
                    var dlg = new CompareRangesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "compare", "diff", "reconcile", "mismatch", "missing values" },
                officeIconName: "CompareDocuments",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "C",
                roadmapBatch: "Active"
            ));

            // 10. Select Special Cells - Module C: Find, Select, and Compare
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SELECT_SPECIAL",
                title: "Select Special Cells",
                description: "Filter and locate cells matching specific attributes (blank, error, formulas, max/min).",
                category: CommandCategory.FindSelectCompare,
                executeAction: ctx =>
                {
                    var dlg = new SelectSpecialDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "select special", "find blanks", "find errors", "max cell", "min cell" },
                officeIconName: "SelectCell",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "C",
                roadmapBatch: "Active"
            ));

            // =========================================================================
            // Module C: Find, Select, and Compare (Batch 4 Expansion)
            // =========================================================================

            // Find Duplicate & Unique Values
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FIND_DUPLICATES",
                title: "Find Duplicates",
                description: "Identify, highlight, or select duplicate or unique values across columns with occurrence rules.",
                category: CommandCategory.FindSelectCompare,
                executeAction: ctx =>
                {
                    var dlg = new FindDuplicatesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "find duplicates", "unique values", "duplicate detector", "select unique", "highlight dups", "distinct values" },
                officeIconName: "FindDialog",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "C",
                roadmapBatch: "Batch 4"
            ));

            // Advanced Multi-Sheet Find & Replace
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_ADVANCED_FIND_REPLACE",
                title: "Multi-Sheet Find",
                description: "Search and batch replace across multiple selected worksheets or entire workbooks with preview.",
                category: CommandCategory.FindSelectCompare,
                executeAction: ctx =>
                {
                    var dlg = new AdvancedFindReplaceDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "advanced find", "multi sheet find", "regex find", "search across sheets", "global replace" },
                officeIconName: "ReplaceDialog",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "C",
                roadmapBatch: "Batch 4"
            ));

            // Select Maximum / Minimum Cells
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SELECT_MAX_MIN",
                title: "Select Max/Min",
                description: "Locate and select the highest or lowest numeric values across selection, rows, or columns.",
                category: CommandCategory.FindSelectCompare,
                executeAction: ctx =>
                {
                    var dlg = new SelectMaxMinDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "select max", "select min", "highest value", "lowest value", "peak cell", "top values", "bottom values" },
                officeIconName: "SelectCell",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "C",
                roadmapBatch: "Batch 4"
            ));

            // Select Row Differences
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SELECT_ROW_DIFFERENCES",
                title: "Row Differences",
                description: "Compare each cell in a row or column against a baseline reference cell and select mismatches.",
                category: CommandCategory.FindSelectCompare,
                executeAction: ctx =>
                {
                    var dlg = new RowDifferencesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "row differences", "column differences", "compare rows", "select differences", "mismatched cells" },
                officeIconName: "CompareDocuments",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "C",
                roadmapBatch: "Batch 4"
            ));

            // Comment Inspector & Manager
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FIND_BY_COMMENT",
                title: "Comment Inspector",
                description: "Inspect, filter, navigate, and audit cell comments across worksheets.",
                category: CommandCategory.FindSelectCompare,
                executeAction: ctx =>
                {
                    var dlg = new CommentInspectorDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "find comments", "select comments", "locate notes", "comment inspector", "audit comments" },
                officeIconName: "ReviewShowOrHideMarkup",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "C",
                roadmapBatch: "Batch 4"
            ));

            // 11. Merge Ranges - Module B: Range and Cell Operations
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_MERGE_RANGES",
                title: "Merge Ranges",
                description: "Combine compatible ranges into one destination by appending rows or side-by-side.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new MergeRangesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "merge", "combine ranges", "append rows", "consolidate", "stack" },
                officeIconName: "TableMergeCells",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Active"
            ));

            // 12. Remove Duplicates - Module E: Data, Filtering, and Statistics
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_DATA_REMOVE_DUPLICATES",
                title: "Remove Duplicates",
                description: "Identify and clear duplicate rows in selected data with key column options.",
                category: CommandCategory.DataAndFiltering,
                executeAction: ctx =>
                {
                    var dlg = new RemoveDuplicatesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "duplicate", "deduplicate", "unique", "distinct" },
                officeIconName: "TableDeleteRows",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "E",
                roadmapBatch: "Active"
            ));

            // 13. Find Missing Records - Module E: Data, Filtering, and Statistics
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FIND_MISSING_RECORDS",
                title: "Find Missing Records",
                description: "Compare two lists by key columns and identify source-only, target-only, or matching records.",
                category: CommandCategory.DataAndFiltering,
                executeAction: ctx =>
                {
                    var dlg = new FindMissingRecordsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "missing records", "reconcile", "difference", "vlookup compare", "unmatched" },
                officeIconName: "FindDialog",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "E",
                roadmapBatch: "Active"
            ));

            // 14. Advanced Filter - Module E: Data, Filtering, and Statistics
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_ADVANCED_FILTER",
                title: "Advanced Filter",
                description: "Filter rows by configurable multi-condition rules (AND/OR) and export or select matches.",
                category: CommandCategory.DataAndFiltering,
                executeAction: ctx =>
                {
                    var dlg = new AdvancedFilterDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "advanced filter", "filter rows", "criteria", "multi condition", "select rows" },
                officeIconName: "Filter",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "E",
                roadmapBatch: "Active"
            ));

            // 15. Convert Formulas to Values - Module G: Formula and Calculation Utilities
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMULA_TO_VALUE",
                title: "Formulas to Values",
                description: "Replace calculated formulas with static values in-place instantly.",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: ctx =>
                {
                    if (ctx.IsCurrentSheetProtected)
                    {
                        ctx.ShowMessage("The active worksheet is protected. Please unprotect the sheet before converting formulas.", "Sheet Protected", isError: true);
                        return;
                    }

                    if (ctx.SelectedCellCount == 0)
                    {
                        ctx.ShowMessage("Please select a range of cells to convert.", "Formulas to Values", isError: true);
                        return;
                    }

                    bool proceed = ctx.Confirm(
                        "Are you sure you want to permanently convert formulas in the selected cells into static values?\n\nThis action replaces formulas with their current calculated values and cannot be automatically undone.",
                        "Confirm Formula Conversion");
                    if (!proceed) return;

                    int affected = ctx.ConvertSelectionFormulasToValues();
                    ctx.ShowMessage($"Converted {affected} cells from formulas to static values.", "Formulas to Values");
                },
                keywords: new[] { "paste values", "freeze formulas", "convert values" },
                officeIconName: "PasteValues",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Active"
            ));

            // 16. Manage Worksheet Visibility - Module D: Worksheet and Workbook Management
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SHEET_VISIBILITY",
                title: "Worksheet Visibility",
                description: "Manage Visible, Hidden, and VeryHidden worksheet states in bulk.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new ManageWorksheetsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "unhide sheets", "hide sheets", "very hidden", "sheet visibility" },
                officeIconName: "ViewNormalViewExcel",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Active"
            ));

            // 17. Batch Rename Worksheets - Module D: Worksheet and Workbook Management
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SHEET_BATCH_RENAME",
                title: "Rename Worksheets",
                description: "Rename multiple sheets simultaneously with prefix, suffix, and numbers.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new BatchRenameWorksheetsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "rename sheets", "bulk rename", "sequence sheets" },
                officeIconName: "SheetRename",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Active"
            ));

            // 18. Split Data to Sheets - Module D: Worksheet and Workbook Management
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SPLIT_DATA_TO_SHEETS",
                title: "Split to Sheets",
                description: "Split dataset into separate worksheets based on distinct values in a key column.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new SplitDataToSheetsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "split to sheets", "separate sheets", "distribute rows", "by column" },
                officeIconName: "TableSplitCells",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Active"
            ));

            // 19. Compare Workbooks - Module D: Worksheet and Workbook Management
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_COMPARE_WORKBOOKS",
                title: "Compare Workbooks",
                description: "Reconcile two open workbooks and generate a difference audit report.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new CompareWorkbooksDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "compare workbooks", "workbook diff", "reconcile workbooks", "audit files" },
                officeIconName: "ViewSideBySide",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Active"
            ));

            // =========================================================================
            // Module D: Worksheet and Workbook Management (Batch 5 Expansion)
            // =========================================================================

            // Combine Worksheets into One
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SHEET_COMBINE",
                title: "Combine Worksheets",
                description: "Consolidate multiple worksheets into a single master sheet with header alignment.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new CombineSheetsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "combine sheets", "merge worksheets", "stack sheets", "master worksheet", "consolidate sheets" },
                officeIconName: "TableMergeCells",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Batch 5"
            ));

            // Combine Workbooks
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_WB_COMBINE_FILES",
                title: "Combine Workbooks",
                description: "Batch import worksheets from multiple external Excel files into the active workbook.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new CombineWorkbooksDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "combine workbooks", "merge excel files", "batch combine", "consolidate files" },
                officeIconName: "FileOpen",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Batch 5"
            ));

            // Split Workbook into Separate Files
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_WB_SPLIT_FILES",
                title: "Split Workbook",
                description: "Save each worksheet as an independent Excel (.xlsx), PDF, or CSV file.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new SplitWorkbookDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "split workbook", "save sheets as files", "separate workbooks", "export sheets", "pdf export" },
                officeIconName: "TableSplitCells",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Batch 5"
            ));

            // Sort Worksheet Tabs
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SHEET_SORT_TABS",
                title: "Sort Worksheet Tabs",
                description: "Reorder worksheet tabs alphabetically (A-Z / Z-A), naturally by numbers, or reversed.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new SortSheetTabsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "sort sheet tabs", "sort worksheets", "alphabetical sheets", "natural sort", "tab order" },
                officeIconName: "SortAscending",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Batch 5"
            ));

            // Create Linked Sheet Index (TOC)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SHEET_CREATE_INDEX",
                title: "Create Sheet Index",
                description: "Generate a Table of Contents worksheet with clickable hyperlinks to every sheet in the workbook.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new CreateSheetIndexDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "sheet index", "table of contents", "linked sheets", "toc generator", "hyperlinks" },
                officeIconName: "HyperlinkInsert",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Batch 5"
            ));

            // Delete Blank Worksheets
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SHEET_DELETE_BLANKS",
                title: "Delete Blank Sheets",
                description: "Scan and batch remove completely empty worksheets after confirmation.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx =>
                {
                    var dlg = new DeleteBlankSheetsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "delete blank sheets", "remove empty worksheets", "clean empty sheets" },
                officeIconName: "SheetDelete",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "D",
                roadmapBatch: "Batch 5"
            ));

            // 20. About Dialog - General / Information
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_ABOUT",
                title: "About Suite",
                description: "View version, architecture, diagnostics, and product information.",
                category: CommandCategory.General,
                executeAction: ctx =>
                {
                    var dlg = new AboutDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "about", "version", "system info", "license" },
                officeIconName: "Info",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "General",
                roadmapBatch: "Active"
            ));

            // Register the comprehensive planned tools across all 12 modules
            RoadmapCatalog.RegisterPlannedTools(registry);
        }
    }
}
