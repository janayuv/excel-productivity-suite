using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum DifferenceDirection
    {
        RowDifferences,    // Compare each cell in a row against the row's baseline column
        ColumnDifferences  // Compare each cell in a column against the column's baseline row
    }

    public class DifferenceOptions
    {
        public DifferenceDirection Direction { get; set; } = DifferenceDirection.RowDifferences;
        public int BaselineIndex { get; set; } = 0; // 0-based column (for row diff) or row (for col diff)
        public bool MatchCase { get; set; } = false;
        public bool TreatBlankAsEmpty { get; set; } = true;
    }

    /// <summary>
    /// Identifies cells whose values differ from a reference comparison cell in the same row or column.
    /// </summary>
    public static class RowDifferenceFinder
    {
        public static List<(int Row, int Column, object? Value, object? BaselineValue)> FindDifferences(
            object?[,] matrix,
            int startRow,
            int startCol,
            DifferenceOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            var stringComparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var differences = new List<(int, int, object?, object?)>();

            if (options.Direction == DifferenceDirection.RowDifferences)
            {
                int baseColOffset = Math.Min(colCount - 1, Math.Max(0, options.BaselineIndex));
                int baseCol = cStart + baseColOffset;

                for (int r = rStart; r <= rEnd; r++)
                {
                    object? baselineVal = matrix[r, baseCol];
                    string baseStr = Normalize(baselineVal, options.TreatBlankAsEmpty);

                    for (int c = cStart; c <= cEnd; c++)
                    {
                        if (c == baseCol) continue; // skip comparison against self

                        object? cellVal = matrix[r, c];
                        string cellStr = Normalize(cellVal, options.TreatBlankAsEmpty);

                        if (!string.Equals(cellStr, baseStr, stringComparison))
                        {
                            int sheetRow = startRow + (r - rStart);
                            int sheetCol = startCol + (c - cStart);
                            differences.Add((sheetRow, sheetCol, cellVal, baselineVal));
                        }
                    }
                }
            }
            else // ColumnDifferences
            {
                int baseRowOffset = Math.Min(rowCount - 1, Math.Max(0, options.BaselineIndex));
                int baseRow = rStart + baseRowOffset;

                for (int c = cStart; c <= cEnd; c++)
                {
                    object? baselineVal = matrix[baseRow, c];
                    string baseStr = Normalize(baselineVal, options.TreatBlankAsEmpty);

                    for (int r = rStart; r <= rEnd; r++)
                    {
                        if (r == baseRow) continue;

                        object? cellVal = matrix[r, c];
                        string cellStr = Normalize(cellVal, options.TreatBlankAsEmpty);

                        if (!string.Equals(cellStr, baseStr, stringComparison))
                        {
                            int sheetRow = startRow + (r - rStart);
                            int sheetCol = startCol + (c - cStart);
                            differences.Add((sheetRow, sheetCol, cellVal, baselineVal));
                        }
                    }
                }
            }

            return differences;
        }

        private static string Normalize(object? val, bool treatBlankAsEmpty)
        {
            if (val == null) return "";
            string s = val.ToString() ?? "";
            return treatBlankAsEmpty ? s.Trim() : s;
        }
    }
}
