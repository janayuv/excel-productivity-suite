using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Worksheets
{
    public class SheetDataInfo
    {
        public string SheetName { get; }
        public object?[,]? Data { get; }

        public SheetDataInfo(string sheetName, object?[,]? data)
        {
            SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
            Data = data;
        }
    }

    public class SheetCombineOptions
    {
        public bool HasHeaders { get; set; } = true;
        public bool KeepFirstSheetHeaderOnly { get; set; } = true;
        public bool AddSourceSheetColumn { get; set; } = true;
        public string SourceColumnHeader { get; set; } = "Source Sheet";
        public int BlankRowsBetweenSheets { get; set; } = 0;
    }

    /// <summary>
    /// Pure computational engine for consolidating multiple 2D worksheet matrices into a master combined dataset.
    /// </summary>
    public static class SheetCombiner
    {
        public static object?[,] Combine(IReadOnlyList<SheetDataInfo> sheets, SheetCombineOptions options, out int totalRowsCombined)
        {
            if (sheets == null) throw new ArgumentNullException(nameof(sheets));
            if (options == null) throw new ArgumentNullException(nameof(options));

            totalRowsCombined = 0;
            if (sheets.Count == 0) return new object?[0, 0];

            int extraCols = options.AddSourceSheetColumn ? 1 : 0;
            int maxCols = 0;
            int totalRows = 0;

            bool isFirstWithData = true;

            // First pass: compute dimension requirements
            foreach (var s in sheets)
            {
                if (s.Data == null) continue;
                int rStart = s.Data.GetLowerBound(0);
                int rEnd = s.Data.GetUpperBound(0);
                int cStart = s.Data.GetLowerBound(1);
                int cEnd = s.Data.GetUpperBound(1);

                int sheetRows = rEnd - rStart + 1;
                int sheetCols = cEnd - cStart + 1;
                if (sheetRows <= 0 || sheetCols <= 0) continue;

                if (sheetCols > maxCols) maxCols = sheetCols;

                if (isFirstWithData)
                {
                    totalRows += sheetRows;
                    isFirstWithData = false;
                }
                else
                {
                    int rowsToAdd = sheetRows;
                    if (options.HasHeaders && options.KeepFirstSheetHeaderOnly && sheetRows > 1)
                    {
                        rowsToAdd -= 1; // omit header
                    }
                    totalRows += rowsToAdd + Math.Max(0, options.BlankRowsBetweenSheets);
                }
            }

            if (totalRows == 0 || maxCols == 0) return new object?[0, 0];

            int finalCols = maxCols + extraCols;
            object?[,] result = new object?[totalRows, finalCols];
            int writeRow = 0;
            bool firstHeaderWritten = false;

            foreach (var s in sheets)
            {
                if (s.Data == null) continue;
                int rStart = s.Data.GetLowerBound(0);
                int rEnd = s.Data.GetUpperBound(0);
                int cStart = s.Data.GetLowerBound(1);
                int cEnd = s.Data.GetUpperBound(1);

                int sheetRows = rEnd - rStart + 1;
                int sheetCols = cEnd - cStart + 1;
                if (sheetRows <= 0 || sheetCols <= 0) continue;

                int readStartRow = rStart;

                // Handle header
                if (!firstHeaderWritten)
                {
                    if (options.AddSourceSheetColumn)
                    {
                        result[writeRow, 0] = options.SourceColumnHeader;
                    }

                    for (int c = 0; c < sheetCols; c++)
                    {
                        result[writeRow, c + extraCols] = s.Data[rStart, cStart + c];
                    }

                    writeRow++;
                    readStartRow = rStart + 1;
                    firstHeaderWritten = true;
                }
                else if (options.HasHeaders && options.KeepFirstSheetHeaderOnly)
                {
                    // Skip header row on subsequent sheets
                    readStartRow = rStart + 1;

                    // Insert blank row spacing if configured
                    writeRow += Math.Max(0, options.BlankRowsBetweenSheets);
                }
                else
                {
                    // Insert blank row spacing if configured
                    writeRow += Math.Max(0, options.BlankRowsBetweenSheets);
                }

                // Copy remaining data rows
                for (int r = readStartRow; r <= rEnd; r++)
                {
                    if (writeRow >= totalRows) break;

                    if (options.AddSourceSheetColumn)
                    {
                        result[writeRow, 0] = s.SheetName;
                    }

                    for (int c = 0; c < sheetCols; c++)
                    {
                        result[writeRow, c + extraCols] = s.Data[r, cStart + c];
                    }

                    writeRow++;
                }
            }

            totalRowsCombined = writeRow;
            return result;
        }
    }
}
