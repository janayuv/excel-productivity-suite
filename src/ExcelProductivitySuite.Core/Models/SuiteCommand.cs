using System;

namespace ExcelProductivitySuite.Core.Models
{
    public enum FeatureStatus
    {
        ImplementedAndVerified,
        ImplementedNeedsVerification,
        PartiallyImplemented,
        Planned,
        Blocked
    }

    public enum CommandCategory
    {
        General = 0,
        TextTools = 1,                  // Module A: Text Tools
        RangeAndCells = 2,              // Module B: Range and Cell Operations
        FindSelectCompare = 3,          // Module C: Find, Select, and Compare
        WorksheetAndWorkbook = 4,       // Module D: Worksheet and Workbook Management
        DataAndFiltering = 5,           // Module E: Data, Filtering, and Statistics
        DataValidation = 6,             // Module F: Data Validation and Dropdown Lists
        FormulasAndCalculation = 7,     // Module G: Formula and Calculation Utilities
        FormattingAndDesign = 8,        // Module H: Formatting and Design
        ImportExportAndFiles = 9,       // Module I: Import, Export, and File Utilities
        ProtectionAndRecovery = 10,     // Module J: Protection and Recovery
        ChartsAndAnalysis = 11,         // Module K: Charts and Analysis
        NavigationAndProductivity = 12, // Module L: Navigation and User Productivity

        // Backward compatibility aliases
        Text = TextTools,
        DataCleaning = DataAndFiltering,
        Worksheet = WorksheetAndWorkbook,
        Workbook = WorksheetAndWorkbook,
        Formulas = FormulasAndCalculation,
        Settings = General
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

        public FeatureStatus Status { get; set; }
        public string ModuleCode { get; set; }
        public string RoadmapBatch { get; set; }

        public bool IsImplemented =>
            Status == FeatureStatus.ImplementedAndVerified ||
            Status == FeatureStatus.ImplementedNeedsVerification;

        public string StatusDisplay => Status switch
        {
            FeatureStatus.ImplementedAndVerified => "Active",
            FeatureStatus.ImplementedNeedsVerification => "Verify",
            FeatureStatus.PartiallyImplemented => "Partial",
            FeatureStatus.Blocked => "Blocked",
            _ => "Roadmap"
        };

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
            string? officeIconName = null,
            FeatureStatus status = FeatureStatus.ImplementedAndVerified,
            string? moduleCode = null,
            string? roadmapBatch = null)
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
            Status = status;
            ModuleCode = moduleCode ?? GetDefaultModuleCode(category);
            RoadmapBatch = roadmapBatch ?? (status == FeatureStatus.ImplementedAndVerified ? "Active" : "Batch 1");
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

        private static string GetDefaultModuleCode(CommandCategory category)
        {
            switch (category)
            {
                case CommandCategory.TextTools: return "A";
                case CommandCategory.RangeAndCells: return "B";
                case CommandCategory.FindSelectCompare: return "C";
                case CommandCategory.WorksheetAndWorkbook: return "D";
                case CommandCategory.DataAndFiltering: return "E";
                case CommandCategory.DataValidation: return "F";
                case CommandCategory.FormulasAndCalculation: return "G";
                case CommandCategory.FormattingAndDesign: return "H";
                case CommandCategory.ImportExportAndFiles: return "I";
                case CommandCategory.ProtectionAndRecovery: return "J";
                case CommandCategory.ChartsAndAnalysis: return "K";
                case CommandCategory.NavigationAndProductivity: return "L";
                default: return "General";
            }
        }
    }
}
