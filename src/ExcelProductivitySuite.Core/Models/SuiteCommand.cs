using System;

namespace ExcelProductivitySuite.Core.Models
{
    public enum CommandCategory
    {
        General,
        Text,
        RangeAndCells,
        DataCleaning,
        Worksheet,
        Workbook,
        Formulas,
        Settings
    }

    /// <summary>
    /// Represents a command available in the Excel Productivity Suite catalog and Ribbon.
    /// </summary>
    public class SuiteCommand
    {
        public string Id { get; }
        public string Title { get; }
        public string Description { get; }
        public CommandCategory Category { get; }
        public string? SubCategory { get; }
        public string[] Keywords { get; }
        public string? ShortcutHint { get; }
        public string? OfficeIconName { get; }
        public bool IsFavorite { get; set; }
        public Action<Abstractions.IExcelContext> ExecuteAction { get; }
        public Func<Abstractions.IExcelContext, bool>? CanExecutePredicate { get; }

        public SuiteCommand(
            string id,
            string title,
            string description,
            CommandCategory category,
            Action<Abstractions.IExcelContext> executeAction,
            Func<Abstractions.IExcelContext, bool>? canExecute = null,
            string? subCategory = null,
            string[]? keywords = null,
            string? shortcutHint = null,
            string? officeIconName = null)
        {
            Id = id ?? throw new ArgumentNullException(nameof(id));
            Title = title ?? throw new ArgumentNullException(nameof(title));
            Description = description ?? string.Empty;
            Category = category;
            ExecuteAction = executeAction ?? throw new ArgumentNullException(nameof(executeAction));
            CanExecutePredicate = canExecute;
            SubCategory = subCategory;
            Keywords = keywords ?? Array.Empty<string>();
            ShortcutHint = shortcutHint;
            OfficeIconName = officeIconName;
        }

        public bool CanExecute(Abstractions.IExcelContext context) =>
            CanExecutePredicate == null || CanExecutePredicate(context);

        public void Execute(Abstractions.IExcelContext context)
        {
            if (CanExecute(context))
            {
                ExecuteAction(context);
            }
        }
    }
}
