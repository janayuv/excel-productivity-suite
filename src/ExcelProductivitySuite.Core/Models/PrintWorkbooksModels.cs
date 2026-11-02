using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Core.Models
{
    public enum PrintFolderMode
    {
        SameDirectory = 0,
        MultipleDirectories = 1
    }

    public enum PrintSheetScope
    {
        EntireWorkbook = 0,
        WorksheetsOnly = 1,
        ChartSheetsOnly = 2,
        WorksheetsAndCharts = 3,
        SpecificSheets = 4
    }

    public enum PrintPageRangeMode
    {
        AllPages = 0,
        FirstPageOnly = 1,
        CustomRange = 2
    }

    public enum PrintItemStatus
    {
        Pending,
        Printing,
        Completed,
        Skipped,
        Failed
    }

    public class WorkbookPrintItem
    {
        public string FilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string Directory { get; set; } = string.Empty;
        public long SizeBytes { get; set; }
        public string FileSizeFormatted { get; set; } = string.Empty;
        public DateTime LastModified { get; set; }
        public bool IsSelected { get; set; } = true;
        public PrintItemStatus Status { get; set; } = PrintItemStatus.Pending;
        public string? ErrorMessage { get; set; }
        public int SheetsPrintedCount { get; set; }

        public static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
            return $"{bytes / (1024.0 * 1024.0):F2} MB";
        }
    }

    public class PrintWorkbooksConfig
    {
        public PrintFolderMode FolderMode { get; set; } = PrintFolderMode.SameDirectory;
        public string DirectoryPath { get; set; } = string.Empty;
        public bool IncludeSubdirectories { get; set; } = false;

        public List<WorkbookPrintItem> SelectedFiles { get; set; } = new List<WorkbookPrintItem>();

        public bool IncludeExcelFiles { get; set; } = true;
        public bool IncludeCsvFiles { get; set; } = false;
        public bool IncludeTextFiles { get; set; } = false;
        public string? CustomExtensionFilter { get; set; }

        public PrintSheetScope SheetScope { get; set; } = PrintSheetScope.EntireWorkbook;
        public List<string> SpecificSheetNames { get; set; } = new List<string>();

        public PrintPageRangeMode PageRangeMode { get; set; } = PrintPageRangeMode.AllPages;
        public int FromPage { get; set; } = 1;
        public int ToPage { get; set; } = 1;
        public int Copies { get; set; } = 1;

        public string? PrinterName { get; set; }
        public bool QuietMode { get; set; } = true;
        public bool DisableMacros { get; set; } = true;
        public bool DoNotUpdateLinks { get; set; } = true;
        public bool GenerateSummaryReport { get; set; } = true;

        public static PrintWorkbooksConfig CreateDefault()
        {
            return new PrintWorkbooksConfig();
        }
    }

    public class PrintWorkbooksResult
    {
        public int TotalWorkbooks { get; set; }
        public int PrintedWorkbooks { get; set; }
        public int FailedWorkbooks { get; set; }
        public int SkippedWorkbooks { get; set; }
        public int TotalSheetsPrinted { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<WorkbookPrintItem> ProcessedItems { get; set; } = new List<WorkbookPrintItem>();
        public TimeSpan Duration { get; set; }
        public string? SummaryReportPath { get; set; }
        public bool IsSuccess => FailedWorkbooks == 0 && Errors.Count == 0;
    }
}
