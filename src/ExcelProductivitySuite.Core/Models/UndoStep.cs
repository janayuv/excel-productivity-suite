using System;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Represents an immutable recorded undoable cell or range operation with workbook, sheet, and coordinates tracking.
    /// </summary>
    public class UndoStep
    {
        public string Description { get; }
        public string WorkbookName { get; }
        public string? WorkbookPath { get; }
        public string SheetName { get; }
        public string RangeAddress { get; }
        public int StartRow { get; }
        public int StartColumn { get; }
        public int RowCount { get; }
        public int ColumnCount { get; }
        public object?[,] BeforeValues { get; }
        public object?[,] AfterValues { get; }
        public object?[,]? BeforeFormulas { get; }
        public object?[,]? AfterFormulas { get; }
        public DateTime Timestamp { get; }

        public UndoStep(
            string description,
            string workbookName,
            string? workbookPath,
            string sheetName,
            string rangeAddress,
            int startRow,
            int startColumn,
            object?[,] beforeValues,
            object?[,] afterValues,
            object?[,]? beforeFormulas = null,
            object?[,]? afterFormulas = null)
        {
            Description = description ?? "Cell Operation";
            WorkbookName = workbookName ?? string.Empty;
            WorkbookPath = workbookPath;
            SheetName = sheetName ?? string.Empty;
            RangeAddress = rangeAddress ?? string.Empty;
            StartRow = startRow;
            StartColumn = startColumn;

            if (beforeValues == null) throw new ArgumentNullException(nameof(beforeValues));
            if (afterValues == null) throw new ArgumentNullException(nameof(afterValues));

            // Defensive deep cloning to guarantee snapshot immutability
            BeforeValues = (object?[,])beforeValues.Clone();
            AfterValues = (object?[,])afterValues.Clone();
            BeforeFormulas = beforeFormulas != null ? (object?[,])beforeFormulas.Clone() : null;
            AfterFormulas = afterFormulas != null ? (object?[,])afterFormulas.Clone() : null;

            RowCount = beforeValues.GetLength(0);
            ColumnCount = beforeValues.GetLength(1);
            Timestamp = DateTime.UtcNow;
        }

        public UndoStep(
            string description,
            string sheetName,
            string rangeAddress,
            int startRow,
            int startColumn,
            object?[,] beforeValues,
            object?[,] afterValues,
            object?[,]? beforeFormulas = null,
            object?[,]? afterFormulas = null)
            : this(description, string.Empty, null, sheetName, rangeAddress, startRow, startColumn, beforeValues, afterValues, beforeFormulas, afterFormulas)
        {
        }

        public override string ToString() => $"{Description} [{WorkbookName}] ({SheetName}!{RangeAddress}) at {Timestamp:HH:mm:ss}";
    }
}
