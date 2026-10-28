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

            // =========================================================================
            // Module E & F: Data, Filtering, and Validation (Batch 6 Expansion)
            // =========================================================================

            // Advanced Multi-Level Sort
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_DATA_ADVANCED_SORT",
                title: "Advanced Sort",
                description: "Sort table by text string length, month names, weekday names, or standard criteria.",
                category: CommandCategory.DataAndFiltering,
                executeAction: ctx =>
                {
                    var dlg = new AdvancedSortDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "advanced sort", "sort by length", "month sort", "weekday sort", "custom sort" },
                officeIconName: "CustomSort",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "E",
                roadmapBatch: "Batch 6"
            ));

            // Filter by Cell / Font Color
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_DATA_FILTER_BY_COLOR",
                title: "Filter by Color",
                description: "Identify and select cells matching a specified background fill or font color.",
                category: CommandCategory.DataAndFiltering,
                executeAction: ctx =>
                {
                    var dlg = new FilterByColorDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "filter by color", "select by color", "fill color", "font color", "color filter" },
                officeIconName: "ColorPicker",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "E",
                roadmapBatch: "Batch 6"
            ));

            // Count and Sum by Color
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_DATA_COUNT_BY_COLOR",
                title: "Count by Color",
                description: "Calculate statistical aggregations (Count, Sum, Average) grouped by cell fill or font color.",
                category: CommandCategory.DataAndFiltering,
                executeAction: ctx =>
                {
                    var dlg = new CountByColorDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "count by color", "sum by color", "average by color", "color statistics", "color aggregation" },
                officeIconName: "FunctionsStatisticalInsertGallery",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "E",
                roadmapBatch: "Batch 6"
            ));

            // Merge / Update Tables on Keys
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_DATA_MERGE_TABLES",
                title: "Merge Tables",
                description: "Upsert and join secondary table records into primary table matching on a common key.",
                category: CommandCategory.DataAndFiltering,
                executeAction: ctx =>
                {
                    var dlg = new MergeTablesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "merge tables", "join tables", "table upsert", "vlookup merge", "combine tables" },
                officeIconName: "Relationship",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "E",
                roadmapBatch: "Batch 6"
            ));

            // Extract Unique / Distinct Records
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_DATA_EXTRACT_DISTINCT",
                title: "Extract Distinct",
                description: "Filter and copy unique records to a new worksheet without modifying source data.",
                category: CommandCategory.DataAndFiltering,
                executeAction: ctx =>
                {
                    var dlg = new ExtractDistinctDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "extract distinct", "unique records", "filter unique", "copy distinct", "deduplicate export" },
                officeIconName: "FilterByResource",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "E",
                roadmapBatch: "Batch 6"
            ));

            // Prevent Duplicate Entries Rule (Data Validation)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_VALIDATION_PREVENT_DUPLICATE",
                title: "Prevent Duplicates",
                description: "Apply dynamic Excel Data Validation to prevent duplicate inputs in the selected range.",
                category: CommandCategory.DataValidation,
                executeAction: ctx =>
                {
                    var dlg = new PreventDuplicateDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "prevent duplicates", "data validation", "unique validation", "reject dups", "prevent duplicate entry" },
                officeIconName: "DataValidation",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "F",
                roadmapBatch: "Batch 6"
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

            // =========================================================================
            // Module G & H: Formulas, Formatting, and Page Setup (Batch 7 Expansion)
            // =========================================================================

            // Convert Formula References (Abs/Rel)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMULA_CONVERT_REF",
                title: "Convert References",
                description: "Batch convert formula cell references between Absolute ($A$1), Row-Absolute (A$1), Column-Absolute ($A1), and Relative (A1).",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: ctx =>
                {
                    var dlg = new ConvertFormulaReferenceDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "convert references", "absolute reference", "relative reference", "dollar sign formula", "lock formula" },
                officeIconName: "FormulaMoreFunctionsDialog",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 7"
            ));

            // Exact Copy Formulas
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMULA_EXACT_COPY",
                title: "Exact Copy Formulas",
                description: "Copy formulas verbatim into new cells or ranges without adjusting relative references.",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: ctx =>
                {
                    var dlg = new ExactCopyFormulaDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "exact copy", "copy formula exact", "duplicate formula", "copy without shifting", "formula copy" },
                officeIconName: "Copy",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 7"
            ));

            // Locate Formula Errors
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMULA_LOCATE_ERRORS",
                title: "Locate Formula Errors",
                description: "Detect, audit, navigate, and highlight formula calculation errors (#DIV/0!, #N/A, #VALUE!, etc.) across the worksheet.",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: ctx =>
                {
                    var dlg = new LocateFormulaErrorsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "locate errors", "find formula errors", "audit errors", "div0", "na error", "error highlighter" },
                officeIconName: "TraceError",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 7"
            ));

            // Math Operations on Selected Cells
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_MATH_OPERATIONS",
                title: "Math Operations",
                description: "Apply arithmetic operations (+, -, *, /, ^, %) directly to cell values in bulk without formula overhead.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new MathOperationsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "math operations", "bulk calculate", "add to cells", "multiply cells", "operation tools", "arithmetic" },
                officeIconName: "Calculator",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 7"
            ));

            // Copy Page Setup Across Sheets
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMAT_PAGE_SETUP_COPY",
                title: "Copy Page Setup",
                description: "Copy page layout, print titles, margins, scaling, and headers/footers to other worksheets.",
                category: CommandCategory.FormattingAndDesign,
                executeAction: ctx =>
                {
                    var dlg = new CopyPageSetupDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "copy page setup", "page setup", "print margins", "repeat header rows", "paper size copy" },
                officeIconName: "PageSetupDialog",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "H",
                roadmapBatch: "Batch 7"
            ));

            // Synchronize Freeze Panes Across Sheets
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMAT_FREEZE_PANES",
                title: "Freeze Panes Multi",
                description: "Apply uniform pane freezing or unfreezing across multiple worksheets simultaneously.",
                category: CommandCategory.FormattingAndDesign,
                executeAction: ctx =>
                {
                    var dlg = new FreezePanesMultiDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "freeze panes", "freeze top row", "freeze multiple sheets", "unfreeze all", "split panes" },
                officeIconName: "FreezePanes",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "H",
                roadmapBatch: "Batch 7"
            ));

            // =========================================================================
            // Module I: Import, Export, and File Utilities (Batch 8 Expansion)
            // =========================================================================

            // Import CSV / Text File
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FILE_IMPORT_CSV",
                title: "Import CSV / Text",
                description: "Parse delimited text and CSV data with automatic formatting and schema preview.",
                category: CommandCategory.ImportExportAndFiles,
                executeAction: ctx =>
                {
                    var dlg = new ImportCsvDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "import csv", "import text", "load delimited file", "read csv into sheet" },
                officeIconName: "GetExternalDataFromText",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "I",
                roadmapBatch: "Batch 8"
            ));

            // Export Range to File
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FILE_EXPORT_RANGE_CSV",
                title: "Export Range to File",
                description: "Export selected cells to CSV, TSV, custom-delimited text, or structured JSON.",
                category: CommandCategory.ImportExportAndFiles,
                executeAction: ctx =>
                {
                    var dlg = new ExportRangeDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "export csv", "save range as csv", "export selection", "export tsv", "export json" },
                officeIconName: "SaveAsCsv",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "I",
                roadmapBatch: "Batch 8"
            ));

            // Export Worksheets to PDF
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FILE_EXPORT_SHEETS_PDF",
                title: "Export to PDF",
                description: "Publish worksheets or entire active workbook as high-resolution PDF documents.",
                category: CommandCategory.ImportExportAndFiles,
                executeAction: ctx =>
                {
                    var dlg = new ExportSheetsPdfDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "export pdf", "save as pdf", "batch pdf export", "convert sheets to pdf" },
                officeIconName: "FileSaveAsPdfOrXps",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "I",
                roadmapBatch: "Batch 8"
            ));

            // Folder File Inventory Listing
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FILE_FOLDER_LISTING",
                title: "Folder File Listing",
                description: "Scan a folder directory and generate a structured inventory of files with hyperlinks.",
                category: CommandCategory.ImportExportAndFiles,
                executeAction: ctx =>
                {
                    var dlg = new FolderListingDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "folder listing", "file inventory", "list files in folder", "directory scan" },
                officeIconName: "OpenFolder",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "I",
                roadmapBatch: "Batch 8"
            ));

            // Insert Text File
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FILE_INSERT_TEXT_FILE",
                title: "Insert Text File",
                description: "Insert raw plain text, log files, or code directly into the active cell or column.",
                category: CommandCategory.ImportExportAndFiles,
                executeAction: ctx =>
                {
                    var dlg = new InsertTextFileDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "insert text file", "append file text", "read text into cell", "load log file" },
                officeIconName: "FileInsert",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "I",
                roadmapBatch: "Batch 8"
            ));

            // Batch Convert Workbooks
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FILE_BATCH_CONVERT",
                title: "Batch Convert Files",
                description: "Convert entire folders of workbook files to modern XLSX, PDF, CSV, or legacy XLS format.",
                category: CommandCategory.ImportExportAndFiles,
                executeAction: ctx =>
                {
                    var dlg = new BatchConvertWorkbooksDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "batch convert workbooks", "convert xls to xlsx", "bulk format conversion", "convert folder" },
                officeIconName: "FileSaveAsExcelXlsx",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "I",
                roadmapBatch: "Batch 8"
            ));

            // =========================================================================
            // Module J: Protection and Recovery (Batch 9 Expansion)
            // =========================================================================

            // Batch Protect Worksheets
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SECURITY_BATCH_PROTECT",
                title: "Batch Protect Sheets",
                description: "Protect multiple worksheets simultaneously with password and granular permissions.",
                category: CommandCategory.ProtectionAndRecovery,
                executeAction: ctx =>
                {
                    var dlg = new BatchProtectSheetsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "batch protect", "protect multiple sheets", "lock worksheets", "sheet security" },
                officeIconName: "ProtectSheet",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "J",
                roadmapBatch: "Batch 9"
            ));

            // Batch Unprotect Worksheets
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SECURITY_BATCH_UNPROTECT",
                title: "Batch Unprotect Sheets",
                description: "Remove worksheet protection across multiple sheets using authorized password.",
                category: CommandCategory.ProtectionAndRecovery,
                executeAction: ctx =>
                {
                    var dlg = new BatchUnprotectSheetsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "batch unprotect", "unlock multiple sheets", "unprotect sheets", "remove password" },
                officeIconName: "UnprotectSheet",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "J",
                roadmapBatch: "Batch 9"
            ));

            // Create Workbook Snapshot
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RECOVERY_CREATE_SNAPSHOT",
                title: "Create Snapshot",
                description: "Capture an instant non-disruptive recovery snapshot of the active workbook.",
                category: CommandCategory.ProtectionAndRecovery,
                executeAction: ctx =>
                {
                    var dlg = new CreateSnapshotDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "workbook snapshot", "take snapshot", "backup state", "undo checkpoint", "restore point" },
                officeIconName: "FileSaveAs",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "J",
                roadmapBatch: "Batch 9"
            ));

            // Restore Workbook Snapshot
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RECOVERY_RESTORE_SNAPSHOT",
                title: "Restore Snapshot",
                description: "Roll back the active workbook to a previously captured safe recovery point.",
                category: CommandCategory.ProtectionAndRecovery,
                executeAction: ctx =>
                {
                    var dlg = new RestoreSnapshotDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "restore snapshot", "revert workbook", "undo to checkpoint", "restore state" },
                officeIconName: "FileHistory",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "J",
                roadmapBatch: "Batch 9"
            ));

            // =========================================================================
            // Module K: Charts and Analysis (Batch 9 Expansion)
            // =========================================================================

            // Quick Statistics Dashboard
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_ANALYSIS_QUICK_STATS",
                title: "Quick Statistics",
                description: "Display instant audit summary with Count, Sum, Mean, Median, Min, Max, and StdDev.",
                category: CommandCategory.ChartsAndAnalysis,
                executeAction: ctx =>
                {
                    var dlg = new QuickStatisticsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "quick stats", "summary statistics", "median mean", "descriptive statistics", "standard deviation" },
                officeIconName: "AutoSum",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "K",
                roadmapBatch: "Batch 9"
            ));

            // Target vs. Actual Chart
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_CHART_TARGET_VS_ACTUAL",
                title: "Target vs. Actual Chart",
                description: "Generate dual-series variance comparison chart (Clustered Column + Target Line).",
                category: CommandCategory.ChartsAndAnalysis,
                executeAction: ctx =>
                {
                    var dlg = new TargetVsActualChartDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "target vs actual", "variance chart", "kpi chart", "performance chart", "column line chart" },
                officeIconName: "ChartInsert",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "K",
                roadmapBatch: "Batch 9"
            ));

            // =========================================================================
            // Module L: Navigation and User Productivity (Batch 10 Expansion)
            // =========================================================================

            // Named Range Manager
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_NAV_NAMED_RANGES",
                title: "Named Range Manager",
                description: "Inspect, audit, filter, navigate, and manage all defined names in the workbook.",
                category: CommandCategory.NavigationAndProductivity,
                executeAction: ctx =>
                {
                    var dlg = new NamedRangeManagerDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "named ranges", "name manager", "locate named range", "inspect names", "broken ref" },
                officeIconName: "NameDefine",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "L",
                roadmapBatch: "Batch 10"
            ));

            // Column Navigator & Manager
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_NAV_COLUMNS",
                title: "Column Navigator",
                description: "Fast navigation, inspection, and visibility toggling across wide worksheets.",
                category: CommandCategory.NavigationAndProductivity,
                executeAction: ctx =>
                {
                    var dlg = new ColumnNavigatorDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "column navigator", "find column", "wide sheet navigation", "hide columns list", "autofit" },
                officeIconName: "AdpColumns",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "L",
                roadmapBatch: "Batch 10"
            ));

            // Cell Content Inspector
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_NAV_CELL_PREVIEW",
                title: "Cell Content Inspector",
                description: "Inspect complete unclipped contents, text metrics, formula expressions, and formatting.",
                category: CommandCategory.NavigationAndProductivity,
                executeAction: ctx =>
                {
                    var dlg = new CellContentInspectorDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "cell inspector", "cell preview", "view long cell", "multiline viewer", "inspect text" },
                officeIconName: "ZoomPrint",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "L",
                roadmapBatch: "Batch 10"
            ));

            // Date Picker Insertion
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_NAV_DATE_PICKER",
                title: "Date Picker",
                description: "Select a calendar date and insert formatted text into the active cell cursor.",
                category: CommandCategory.NavigationAndProductivity,
                executeAction: ctx =>
                {
                    var dlg = new DatePickerDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "date picker", "calendar picker", "insert date popup", "cell calendar", "iso date" },
                officeIconName: "AppointmentColorDialog",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "L",
                roadmapBatch: "Batch 10"
            ));

            // Reading Grid Focus Crosshair
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_NAV_GRID_FOCUS",
                title: "Reading Grid Focus",
                description: "Row & column crosshair tracking assistant for effortlessly navigating wide tables.",
                category: CommandCategory.NavigationAndProductivity,
                executeAction: ctx =>
                {
                    var dlg = new ReadingGridFocusDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "grid focus", "crosshair highlighting", "reading mode", "highlight active row", "table navigation" },
                officeIconName: "TableSelectRowAndColumn",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "L",
                roadmapBatch: "Batch 10"
            ));

            // Table Data Entry Form
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_NAV_DATA_ENTRY_FORM",
                title: "Table Data Entry Form",
                description: "Dynamic structured data entry and record navigation for active worksheet tables.",
                category: CommandCategory.NavigationAndProductivity,
                executeAction: ctx =>
                {
                    var dlg = new DataEntryFormDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "data entry form", "record form", "add row dialog", "structured input", "table form" },
                officeIconName: "FormNew",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "L",
                roadmapBatch: "Batch 10"
            ));

            // =========================================================================
            // Module G & H: Formulas & Formatting Expansion (Batch 11)
            // =========================================================================

            // Audit Formula Precedents
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMULA_AUDIT_PRECEDENTS",
                title: "Audit Formula Precedents",
                description: "Inspect and navigate cell references feeding into the active formula.",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: ctx =>
                {
                    var dlg = new FormulaAuditPrecedentsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "audit precedents", "trace precedents", "precedent cells", "formula dependencies" },
                officeIconName: "TracePrecedents",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 11"
            ));

            // Audit Formula Dependents
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMULA_AUDIT_DEPENDENTS",
                title: "Audit Formula Dependents",
                description: "Inspect and navigate downstream cells that depend on the active cell value.",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: ctx =>
                {
                    var dlg = new FormulaAuditDependentsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "audit dependents", "trace dependents", "dependent cells", "downstream calculations" },
                officeIconName: "TraceDependents",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 11"
            ));

            // Replace Formula References
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMULA_BATCH_REPLACE_REF",
                title: "Replace Formula References",
                description: "Search and substitute sheet references, ranges, or functions across formulas.",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: ctx =>
                {
                    var dlg = new FormulaBatchReplaceDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "replace references", "formula find replace", "rename sheet in formulas", "batch formula edit" },
                officeIconName: "ReplaceDialog",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 11"
            ));

            // Formula Templates Library
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMULA_TEMPLATES",
                title: "Formula Templates Library",
                description: "Insert production-tested formulas for financial models, lookups, math, and date calculations.",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: ctx =>
                {
                    var dlg = new FormulaTemplatesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "formula templates", "financial formulas", "cagr", "safe vlookup", "date difference", "two way lookup" },
                officeIconName: "FunctionsFinancialInsertGallery",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 11"
            ));

            // Batch Number Formatting
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMAT_BATCH_NUMBERS",
                title: "Batch Number Formatting",
                description: "Apply standardized financial, decimal, percentage, and currency masks across ranges.",
                category: CommandCategory.FormattingAndDesign,
                executeAction: ctx =>
                {
                    var dlg = new BatchNumberFormatDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "number format", "currency format", "accounting format", "decimal places", "percentage format" },
                officeIconName: "FormatNumberCurrency",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "H",
                roadmapBatch: "Batch 11"
            ));

            // Standardize Date Formats
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMAT_DATE_STANDARDS",
                title: "Standardize Date Formats",
                description: "Format date cells to ISO 8601, European, US, or enterprise standards in bulk.",
                category: CommandCategory.FormattingAndDesign,
                executeAction: ctx =>
                {
                    var dlg = new StandardizeDateFormatsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "date format", "iso 8601", "standardize dates", "convert text dates", "date masks" },
                officeIconName: "AppointmentColorDialog",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "H",
                roadmapBatch: "Batch 11"
            ));

            // =========================================================================
            // Module B & C: Advanced Range Operations & Copy/Paste Expansion (Batch 12)
            // =========================================================================

            // CMD_RANGE_COPY_NONCONTIGUOUS
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_COPY_NONCONTIGUOUS",
                title: "Copy Non-Contiguous Selection",
                description: "Copy multiple disjoint selection areas into a single contiguous destination range.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new CopyMultipleRangesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "copy noncontiguous", "disjoint copy", "multi area copy", "ctrl select copy", "stack ranges" },
                officeIconName: "Copy",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 12"
            ));

            // CMD_RANGE_PASTE_VISIBLE_ONLY
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_PASTE_VISIBLE_ONLY",
                title: "Paste to Visible Cells Only",
                description: "Paste clipboard data strictly into visible cells, bypassing filtered or hidden rows.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new PasteVisibleOnlyDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "paste visible", "skip hidden rows", "paste filtered", "visible cells only" },
                officeIconName: "Paste",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 12"
            ));

            // CMD_RANGE_UNMERGE_FILL
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_UNMERGE_FILL",
                title: "Unmerge and Fill Cells",
                description: "Safely unmerge composite cells and automatically populate all constituent cells with the original value.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new UnmergeAndFillDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "unmerge and fill", "split merged", "fill unmerged", "normalize merged cells" },
                officeIconName: "TableSplitCells",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 12"
            ));

            // CMD_RANGE_RANDOM_SAMPLE
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_RANDOM_SAMPLE",
                title: "Random Sampling of Records",
                description: "Extract an unbiased random sample of N rows or P% of records from a dataset (with/without replacement, seedable for audits).",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx =>
                {
                    var dlg = new RandomSampleDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "random sample", "sample rows", "extract sample", "audit sample", "statistical sampling" },
                officeIconName: "FunctionsStatisticalInsertGallery",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Batch 12"
            ));

            // CMD_FIND_BY_FORMAT
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FIND_BY_FORMAT",
                title: "Find Cells by Format & Color",
                description: "Locate and select cells matching specific font color, fill color, font style, or border attributes.",
                category: CommandCategory.FindSelectCompare,
                executeAction: ctx =>
                {
                    var dlg = new FindByFormatDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "find by color", "find formatting", "match cell color", "find bold cells", "select formatted" },
                officeIconName: "ColorPicker",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "C",
                roadmapBatch: "Batch 12"
            ));

            // CMD_FIND_BY_VALIDATION
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FIND_BY_VALIDATION",
                title: "Find Cells with Data Validation",
                description: "Audit, inspect, filter, and select cells governed by Excel data validation rules.",
                category: CommandCategory.FindSelectCompare,
                executeAction: ctx =>
                {
                    var dlg = new FindByValidationDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "find validation", "cells with dropdowns", "validation cells", "inspect validation", "audit dropdown" },
                officeIconName: "DataValidation",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "C",
                roadmapBatch: "Batch 12"
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
