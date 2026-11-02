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
            // Module D: Worksheet and Workbook Management (Remaining Future Roadmap)
            // =========================================================================
            AddPlanned(registry, "CMD_SHEET_CONSOLIDATE", "Consolidate Data Across Sheets",
                "Aggregate tabular data across identical structured sheets using SUM, AVERAGE, or COUNT.",
                CommandCategory.WorksheetAndWorkbook, "D", "Batch 5",
                new[] { "consolidate sheets", "3d consolidation", "sum across sheets", "multi sheet rollup" }, "Consolidate");

            // =========================================================================
            // Module E: Data, Filtering, and Statistics (Remaining Future Roadmap)
            // =========================================================================
            AddPlanned(registry, "CMD_DATA_QUICK_SUBTOTAL", "Generate Summaries & Subtotals",
                "Insert grouping subtotals and summaries at every change in designated categorical columns.",
                CommandCategory.DataAndFiltering, "E", "Batch 6+",
                new[] { "subtotals", "group summaries", "data subtotals", "category rollup" }, "PivotTableSubtotals");

            // =========================================================================
            // Module F: Data Validation and Dropdown Lists (Remaining Future Roadmap)
            // =========================================================================
            AddPlanned(registry, "CMD_VALIDATION_SEARCHABLE_DROPDOWN", "Searchable Dropdown Interface",
                "Provide an enhanced searchable popup list for fast data entry in cells with large validation lists.",
                CommandCategory.DataValidation, "F", "Batch 6+",
                new[] { "searchable dropdown", "autocomplete dropdown", "fast validation picker", "filter list" }, "FindDialog");

            AddPlanned(registry, "CMD_VALIDATION_MULTI_SELECT", "Multi-Selection Dropdown List",
                "Allow selecting multiple items from a dropdown list separated by commas within a single cell.",
                CommandCategory.DataValidation, "F", "Batch 6+",
                new[] { "multi select dropdown", "multiple choices", "comma separated dropdown", "multi pick" }, "CheckMark");

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

            AddPlanned(registry, "CMD_VALIDATION_COPY_RULES", "Copy Validation Rules",
                "Copy data validation specifications from a source range to multiple target ranges.",
                CommandCategory.DataValidation, "F", "Batch 6",
                new[] { "copy validation", "duplicate validation rules", "sync dropdowns" }, "Copy");

            // =========================================================================
            // Module G: Formula and Calculation Utilities (Remaining Future Roadmap)
            // =========================================================================
            AddPlanned(registry, "CMD_FORMULA_EXPLAIN", "Explain Formula Deterministically",
                "Deconstruct and explain nested formulas in plain English using a local deterministic parser.",
                CommandCategory.FormulasAndCalculation, "G", "Batch 12+",
                new[] { "explain formula", "formula parser", "deconstruct formula", "understand formula" }, "Info");

            // =========================================================================
            // Module H: Formatting and Design (Remaining Future Roadmap)
            // =========================================================================
            AddPlanned(registry, "CMD_FORMAT_HIGHLIGHT_CELL_TYPES", "Highlight Cell Types",
                "Apply distinct background colors to formulas, constants, locked cells, errors, and text.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "highlight cell types", "color formulas", "color constants", "visual cell audit" }, "ColorPicker");

            AddPlanned(registry, "CMD_FORMAT_BORDERS", "Configurable Border Schemes",
                "Apply elegant table borders, gridlines, thick headers, or double bottom totals.",
                CommandCategory.FormattingAndDesign, "H", "Batch 7",
                new[] { "table borders", "border styles", "accounting borders", "gridlines" }, "BorderAll");

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
                CommandCategory.FormattingAndDesign, "H", "Batch 7+",
                new[] { "repeating headers", "print titles", "rows to repeat at top", "print header rows" }, "PrintTitles");

            // =========================================================================
            // Module I: Import, Export, and File Utilities (Remaining Future Roadmap)
            // =========================================================================
            AddPlanned(registry, "CMD_FILE_IMPORT_PICTURES", "Import Pictures into Cells",
                "Batch import image files into cells with automated sizing and cell locking.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8+",
                new[] { "import pictures", "insert images in cells", "batch photos", "fit image to cell" }, "PictureInsertFromFile");

            AddPlanned(registry, "CMD_FILE_MATCH_PICTURES", "Match Pictures by Key",
                "Match image file names against product codes or employee IDs in cells and insert pictures.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8+",
                new[] { "match pictures", "lookup images", "insert pictures by name", "product photos" }, "PictureInsertFromFile");

            AddPlanned(registry, "CMD_FILE_EXPORT_CHARTS", "Export Charts to Images",
                "Export all embedded charts and shapes in the workbook as PNG, JPEG, or SVG files.",
                CommandCategory.ImportExportAndFiles, "I", "Batch 8+",
                new[] { "export charts", "save chart as image", "chart picture export", "png charts" }, "ChartTypePieInsertGallery");

            // =========================================================================
            // Module J: Protection and Recovery (Batch 9 Expansion)
            // =========================================================================
            AddPlanned(registry, "CMD_SECURITY_WORKBOOK_ENCRYPT", "Manage Workbook Password",
                "Securely set or clear workbook open and modify passwords.",
                CommandCategory.ProtectionAndRecovery, "J", "Batch 9+",
                new[] { "workbook password", "encrypt workbook", "protect workbook structure" }, "ProtectWorkbook");

            AddPlanned(registry, "CMD_SECURITY_WINDOW_VISIBILITY", "Manage Workbook Windows",
                "Batch hide or unhide workbook windows across multi-document sessions.",
                CommandCategory.ProtectionAndRecovery, "J", "Batch 9+",
                new[] { "hide windows", "unhide window", "window visibility", "excel windows" }, "WindowHide");

            // =========================================================================
            // Module K: Charts and Analysis (Batch 9 Expansion)
            // =========================================================================
            AddPlanned(registry, "CMD_CHART_TEMPLATES", "Reusable Chart Templates",
                "Apply standardized corporate color palettes and clean formatting to selected charts.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9+",
                new[] { "chart templates", "format chart", "corporate charts", "clean chart style" }, "ChartTypeColumnInsertGallery");

            AddPlanned(registry, "CMD_CHART_GANTT", "Generate Project Gantt Chart",
                "Create a project timeline Gantt chart from task names, start dates, and durations.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9+",
                new[] { "gantt chart", "project timeline", "schedule chart", "task bar chart" }, "ChartTypeBarInsertGallery");

            AddPlanned(registry, "CMD_CHART_WATERFALL", "Generate Waterfall Chart",
                "Construct a bridge / waterfall variance chart compatible with Excel 2016.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9+",
                new[] { "waterfall chart", "bridge chart", "variance analysis", "financial chart" }, "ChartTypeColumnInsertGallery");

            AddPlanned(registry, "CMD_ANALYSIS_PIVOT_HELPERS", "PivotTable Utilities",
                "Bulk refresh all PivotTables, adjust field formatting, and toggle classic layout.",
                CommandCategory.ChartsAndAnalysis, "K", "Batch 9",
                new[] { "pivot table helpers", "refresh pivot tables", "pivot formatting", "classic pivot" }, "PivotTableInsert");

            // =========================================================================
            // Module L: Navigation and User Productivity (Batch 10 Expansion)
            // =========================================================================

            AddPlanned(registry, "CMD_NAV_FORMULA_EDITOR", "Expanded Multi-Line Formula Editor",
                "Edit long, complex nested formulas with syntax highlighting, indentation, and formatting.",
                CommandCategory.NavigationAndProductivity, "L", "Batch 10+",
                new[] { "expanded formula editor", "multiline formula", "big formula editor", "format formula" }, "FormulaMoreFunctionsMenu");
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
