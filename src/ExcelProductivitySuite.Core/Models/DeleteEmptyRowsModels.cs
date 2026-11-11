using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Defines the target scope for empty row deletion operations.
    /// Corresponds to ASAP Utilities Tool #232 (Delete All Empty Rows).
    /// </summary>
    public enum DeleteEmptyRowsScope
    {
        /// <summary>
        /// Entire active worksheet dataset (UsedRange). Matches ASAP Utilities default behavior.
        /// </summary>
        ActiveWorksheet,

        /// <summary>
        /// Only within currently selected cells or rows.
        /// </summary>
        CurrentSelection,

        /// <summary>
        /// All worksheets in the active workbook.
        /// </summary>
        AllWorksheetsInWorkbook,

        /// <summary>
        /// Explicit list of worksheet names.
        /// </summary>
        SelectedWorksheets
    }

    /// <summary>
    /// Configuration options for scanning and deleting completely empty rows.
    /// </summary>
    public class DeleteEmptyRowsOptions
    {
        /// <summary>
        /// Target scope of rows to scan and delete.
        /// </summary>
        public DeleteEmptyRowsScope Scope { get; set; } = DeleteEmptyRowsScope.ActiveWorksheet;

        /// <summary>
        /// Whether cells containing only whitespace characters (spaces, tabs, newlines) are treated as empty.
        /// Default is true.
        /// </summary>
        public bool TreatWhitespaceAsEmpty { get; set; } = true;

        /// <summary>
        /// Whether formula cells that evaluate to an empty string ("") or null are treated as empty.
        /// Default is true.
        /// </summary>
        public bool TreatEmptyFormulaAsEmpty { get; set; } = true;

        /// <summary>
        /// Whether to skip rows that are hidden or filtered out.
        /// Default is false (deletes empty rows regardless of visibility).
        /// </summary>
        public bool SkipHiddenRows { get; set; } = false;

        /// <summary>
        /// Prompt user with confirmation dialog if multiple rows are found.
        /// </summary>
        public bool ConfirmBeforeDelete { get; set; } = true;

        /// <summary>
        /// Target worksheet names when Scope is SelectedWorksheets.
        /// </summary>
        public List<string>? TargetWorksheets { get; set; }
    }

    /// <summary>
    /// Result of scanning a worksheet or selection for empty rows prior to deletion.
    /// </summary>
    public class EmptyRowsScanResult
    {
        public string WorksheetName { get; set; } = string.Empty;
        public int TotalRowsScanned { get; set; }
        public int StartRow { get; set; } = 1;
        public int EndRow { get; set; } = 1;
        public int StartCol { get; set; } = 1;
        public int EndCol { get; set; } = 1;

        /// <summary>
        /// 1-based worksheet row numbers that are completely empty.
        /// </summary>
        public List<int> EmptyRowIndices { get; set; } = new List<int>();

        /// <summary>
        /// Contiguous ranges of empty rows (startRow, endRow inclusive) for batch deletion.
        /// </summary>
        public List<(int startRow, int endRow)> EmptyRowRanges { get; set; } = new List<(int, int)>();

        public int EmptyRowCount => EmptyRowIndices.Count;
        public int RemainingRowCount => Math.Max(0, TotalRowsScanned - EmptyRowCount);

        public string FormattedRowRanges => FormatRanges(EmptyRowRanges);

        public static string FormatRanges(List<(int startRow, int endRow)> ranges)
        {
            if (ranges == null || ranges.Count == 0) return "None";
            var parts = new List<string>();
            foreach (var r in ranges)
            {
                if (r.startRow == r.endRow)
                    parts.Add(r.startRow.ToString());
                else
                    parts.Add($"{r.startRow}-{r.endRow}");
            }
            return string.Join(", ", parts);
        }
    }

    /// <summary>
    /// Outcome of a batch delete empty rows execution.
    /// </summary>
    public class DeleteEmptyRowsResult
    {
        public bool Success { get; set; }
        public int TotalRowsDeleted { get; set; }
        public Dictionary<string, int> DeletedRowsPerSheet { get; set; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public string? ErrorMessage { get; set; }
        public string Message { get; set; } = string.Empty;

        public static DeleteEmptyRowsResult Succeeded(int count, string message, Dictionary<string, int>? sheetDetails = null)
        {
            return new DeleteEmptyRowsResult
            {
                Success = true,
                TotalRowsDeleted = count,
                Message = message,
                DeletedRowsPerSheet = sheetDetails ?? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            };
        }

        public static DeleteEmptyRowsResult Failed(string error)
        {
            return new DeleteEmptyRowsResult
            {
                Success = false,
                TotalRowsDeleted = 0,
                ErrorMessage = error,
                Message = error
            };
        }
    }
}
