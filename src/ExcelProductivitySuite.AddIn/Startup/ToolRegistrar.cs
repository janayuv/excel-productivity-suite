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

            // 1. Tool Launcher (SuperSearch)
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_LAUNCHER",
                title: "Tool Launcher",
                description: "Search, filter, and launch any productivity utility instantly.",
                category: CommandCategory.General,
                executeAction: ctx =>
                {
                    var dlg = new ToolLauncherWindow(registry, ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "search", "launcher", "find tool", "palette", "commands" },
                shortcutHint: "Ctrl+Shift+P",
                officeIconName: "FindDialog"
            ));

            // 2. Extract Text
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_EXTRACT",
                title: "Extract Text",
                description: "Extract substrings before, after, or between delimiters into cells.",
                category: CommandCategory.Text,
                executeAction: ctx =>
                {
                    var dlg = new ExtractTextDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "extract", "substring", "delimiter", "parse", "slice" },
                officeIconName: "Cut"
            ));

            // 3. Split Text
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_SPLIT",
                title: "Split Cells",
                description: "Split delimited cell values across multiple columns or into new rows.",
                category: CommandCategory.Text,
                executeAction: ctx =>
                {
                    var dlg = new SplitTextDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "split", "columns", "rows", "delimiter", "tokenize" },
                officeIconName: "TableSplitCells"
            ));

            // 4. Remove Characters
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_REMOVE_CHARS",
                title: "Remove Characters",
                description: "Strip numeric, alphabetic, symbols, or custom character sets.",
                category: CommandCategory.Text,
                executeAction: ctx =>
                {
                    var dlg = new RemoveCharactersDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "remove", "delete characters", "strip digits", "clean symbols" },
                officeIconName: "Delete"
            ));

            // 5. Trim & Normalize Whitespace
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_TRIM_SPACES",
                title: "Trim Spaces",
                description: "Clean excess spaces, leading/trailing gaps, and non-breaking spaces.",
                category: CommandCategory.Text,
                executeAction: ctx =>
                {
                    var dlg = new WhitespaceDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "trim", "whitespace", "spaces", "clean", "nbsp" },
                officeIconName: "ParagraphSpacingAfter"
            ));

            // 6. Change Text Case
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_CHANGE_CASE",
                title: "Change Case",
                description: "Convert text to UPPERCASE, lowercase, Proper Case, or Sentence case.",
                category: CommandCategory.Text,
                executeAction: ctx =>
                {
                    var dlg = new ChangeCaseDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "case", "uppercase", "lowercase", "proper case", "capitalize" },
                officeIconName: "ChangeCase"
            ));

            // 7. Add Text / Affixes
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_ADD_AFFIX",
                title: "Add Text",
                description: "Prepend prefix or append suffix to cells with duplicate safeguards.",
                category: CommandCategory.Text,
                executeAction: ctx =>
                {
                    var dlg = new AddTextDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "prefix", "suffix", "append", "prepend", "affix" },
                officeIconName: "AdpAddExisting"
            ));

            // 8. Fill Blank Cells
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
                officeIconName: "FillDown"
            ));

            // 9. Remove Duplicates
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_DATA_REMOVE_DUPLICATES",
                title: "Remove Duplicates",
                description: "Identify and clear duplicate rows in selected data.",
                category: CommandCategory.DataCleaning,
                executeAction: ctx =>
                {
                    var dlg = new RemoveDuplicatesDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "duplicate", "deduplicate", "unique", "distinct" },
                officeIconName: "TableDeleteRows"
            ));

            // 10. Convert Formulas to Values
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_FORMULA_TO_VALUE",
                title: "Formulas to Values",
                description: "Replace calculated formulas with static values in-place instantly.",
                category: CommandCategory.DataCleaning,
                executeAction: ctx =>
                {
                    int affected = ctx.ConvertSelectionFormulasToValues();
                    ctx.ShowMessage($"Converted {affected} cells from formulas to static values.", "Formulas to Values");
                },
                keywords: new[] { "paste values", "freeze formulas", "convert values" },
                officeIconName: "PasteValues"
            ));

            // 11. Manage Worksheet Visibility
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SHEET_VISIBILITY",
                title: "Worksheet Visibility",
                description: "Manage Visible, Hidden, and VeryHidden worksheet states in bulk.",
                category: CommandCategory.Worksheet,
                executeAction: ctx =>
                {
                    var dlg = new ManageWorksheetsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "unhide sheets", "hide sheets", "very hidden", "sheet visibility" },
                officeIconName: "ViewNormalViewExcel"
            ));

            // 12. Batch Rename Worksheets
            registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SHEET_BATCH_RENAME",
                title: "Rename Worksheets",
                description: "Rename multiple sheets simultaneously with prefix, suffix, and numbers.",
                category: CommandCategory.Worksheet,
                executeAction: ctx =>
                {
                    var dlg = new BatchRenameWorksheetsDialog(ctx);
                    dlg.ShowModal(ctx);
                },
                keywords: new[] { "rename sheets", "bulk rename", "sequence sheets" },
                officeIconName: "SheetRename"
            ));

            // 13. About Dialog
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
                officeIconName: "Info"
            ));
        }
    }
}
