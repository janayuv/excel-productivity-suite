using System;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Represents the visibility state of an Excel worksheet.
    /// </summary>
    public enum SheetVisibilityState
    {
        Visible = 0,
        Hidden = 1,
        VeryHidden = 2
    }

    /// <summary>
    /// Sorting order for open workbooks in the navigation list.
    /// </summary>
    public enum WorkbookSortOrder
    {
        OpenOrder = 0,
        NameAscending = 1,
        NameDescending = 2
    }

    /// <summary>
    /// Filter mode for the worksheet navigation list.
    /// </summary>
    public enum SheetFilterMode
    {
        All = 0,
        VisibleOnly = 1,
        HiddenOnly = 2
    }

    /// <summary>
    /// Direction for moving/reordering worksheets in a workbook.
    /// </summary>
    public enum SheetMoveDirection
    {
        Up = 0,
        Down = 1,
        Top = 2,
        End = 3
    }

    /// <summary>
    /// Snapshot summary of an open Excel workbook.
    /// </summary>
    public class WorkbookSummaryInfo
    {
        public string Name { get; }
        public string Path { get; }
        public bool IsActive { get; }
        public bool IsSaved { get; }
        public int SheetCount { get; }
        public bool IsReadOnly { get; }

        public string DisplayName => IsSaved ? Name : $"{Name} *";
        public string StatusDisplay => IsSaved ? "Saved" : "Modified";
        public bool HasDiskPath => !string.IsNullOrWhiteSpace(Path) && !Path.Equals("[Unsaved]", StringComparison.OrdinalIgnoreCase);

        public WorkbookSummaryInfo(
            string name,
            string? path,
            bool isActive,
            bool isSaved,
            int sheetCount,
            bool isReadOnly = false)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Path = string.IsNullOrWhiteSpace(path) ? "[Unsaved]" : path!;
            IsActive = isActive;
            IsSaved = isSaved;
            SheetCount = sheetCount >= 0 ? sheetCount : 0;
            IsReadOnly = isReadOnly;
        }
    }

    /// <summary>
    /// Snapshot summary of an individual worksheet within an open workbook.
    /// </summary>
    public class WorksheetSummaryInfo
    {
        public string Name { get; }
        public string WorkbookName { get; }
        public int Index { get; }
        public bool IsActive { get; }
        public SheetVisibilityState Visibility { get; }
        public int? TabColorRgb { get; }
        public bool IsProtected { get; }

        public bool IsVisible => Visibility == SheetVisibilityState.Visible;
        public bool IsHidden => Visibility != SheetVisibilityState.Visible;

        public string VisibilityDisplay => Visibility switch
        {
            SheetVisibilityState.Visible => "Visible",
            SheetVisibilityState.Hidden => "Hidden",
            SheetVisibilityState.VeryHidden => "Very Hidden",
            _ => "Visible"
        };

        public WorksheetSummaryInfo(
            string name,
            string workbookName,
            int index,
            bool isActive,
            SheetVisibilityState visibility,
            int? tabColorRgb = null,
            bool isProtected = false)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            WorkbookName = workbookName ?? throw new ArgumentNullException(nameof(workbookName));
            Index = index;
            IsActive = isActive;
            Visibility = visibility;
            TabColorRgb = tabColorRgb;
            IsProtected = isProtected;
        }
    }
}
