using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum BlankRowInsertMode
    {
        EveryNRows,          // Insert blank rows at regular intervals
        OnValueChange        // Insert blank rows when key column value changes
    }

    public class BlankRowInsertOptions
    {
        public BlankRowInsertMode Mode { get; set; } = BlankRowInsertMode.EveryNRows;
        public int Interval { get; set; } = 1;             // Insert every N rows
        public int BlankRowCount { get; set; } = 1;        // Number of blank rows to insert
        public int KeyColumnIndex { get; set; } = 0;       // 0-based column index for value change
        public bool HasHeaderRow { get; set; } = true;
    }

    public class BlankRowDeleteOptions
    {
        public bool CheckEntireRow { get; set; } = true;   // All cells in row must be blank
        public int KeyColumnIndex { get; set; } = 0;       // Or only key column must be blank
        public bool DeleteBlankColumns { get; set; } = false; // Delete blank columns instead
    }

    /// <summary>
    /// Inserts or deletes blank rows and columns based on intervals, value transitions, or emptiness.
    /// </summary>
    public static class BlankRowManager
    {
        public static object?[,] InsertBlankRows(object?[,] source, BlankRowInsertOptions options)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            if (rowCount == 0 || colCount == 0) return source;

            var newRows = new List<object?[]>();

            int startRowIdx = 0;
            if (options.HasHeaderRow && rowCount > 0)
            {
                var header = new object?[colCount];
                for (int c = 0; c < colCount; c++) header[c] = source[rStart, cStart + c];
                newRows.Add(header);
                startRowIdx = 1;
            }

            int interval = Math.Max(1, options.Interval);
            int blankCount = Math.Max(1, options.BlankRowCount);
            object? prevVal = null;
            bool isFirstDataRow = true;

            int rowsSinceLastBlank = 0;

            for (int r = startRowIdx; r < rowCount; r++)
            {
                var rowData = new object?[colCount];
                for (int c = 0; c < colCount; c++) rowData[c] = source[rStart + r, cStart + c];

                if (options.Mode == BlankRowInsertMode.OnValueChange)
                {
                    int keyCol = Math.Min(colCount - 1, Math.Max(0, options.KeyColumnIndex));
                    object? currentVal = rowData[keyCol];

                    if (!isFirstDataRow && !Equals(prevVal, currentVal))
                    {
                        for (int b = 0; b < blankCount; b++)
                        {
                            newRows.Add(new object?[colCount]);
                        }
                    }

                    prevVal = currentVal;
                    isFirstDataRow = false;
                    newRows.Add(rowData);
                }
                else // EveryNRows
                {
                    newRows.Add(rowData);
                    rowsSinceLastBlank++;

                    if (rowsSinceLastBlank >= interval && r < rowCount - 1)
                    {
                        for (int b = 0; b < blankCount; b++)
                        {
                            newRows.Add(new object?[colCount]);
                        }
                        rowsSinceLastBlank = 0;
                    }
                }
            }

            // Construct new 2D array
            object?[,] result = new object?[newRows.Count, colCount];
            for (int r = 0; r < newRows.Count; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    result[r, c] = newRows[r][c];
                }
            }

            return result;
        }

        public static object?[,] DeleteBlankRows(object?[,] source, BlankRowDeleteOptions options, out int deletedCount)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            if (options.DeleteBlankColumns)
            {
                var keepCols = new List<int>();
                for (int c = 0; c < colCount; c++)
                {
                    bool isBlank = true;
                    for (int r = 0; r < rowCount; r++)
                    {
                        var val = source[rStart + r, cStart + c];
                        if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                        {
                            isBlank = false;
                            break;
                        }
                    }
                    if (!isBlank) keepCols.Add(c);
                }

                deletedCount = colCount - keepCols.Count;
                if (keepCols.Count == 0) return new object?[rowCount, 0];

                object?[,] colResult = new object?[rowCount, keepCols.Count];
                for (int r = 0; r < rowCount; r++)
                {
                    for (int i = 0; i < keepCols.Count; i++)
                    {
                        colResult[r, i] = source[rStart + r, cStart + keepCols[i]];
                    }
                }
                return colResult;
            }

            var keepRows = new List<int>();
            for (int r = 0; r < rowCount; r++)
            {
                bool isBlank;
                if (options.CheckEntireRow)
                {
                    isBlank = true;
                    for (int c = 0; c < colCount; c++)
                    {
                        var val = source[rStart + r, cStart + c];
                        if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                        {
                            isBlank = false;
                            break;
                        }
                    }
                }
                else
                {
                    int keyCol = Math.Min(colCount - 1, Math.Max(0, options.KeyColumnIndex));
                    var val = source[rStart + r, cStart + keyCol];
                    isBlank = val == null || string.IsNullOrWhiteSpace(val.ToString());
                }

                if (!isBlank)
                {
                    keepRows.Add(r);
                }
            }

            deletedCount = rowCount - keepRows.Count;
            if (keepRows.Count == 0) return new object?[0, colCount];

            object?[,] result = new object?[keepRows.Count, colCount];
            for (int i = 0; i < keepRows.Count; i++)
            {
                int srcRow = keepRows[i];
                for (int c = 0; c < colCount; c++)
                {
                    result[i, c] = source[rStart + srcRow, cStart + c];
                }
            }

            return result;
        }
    }
}
