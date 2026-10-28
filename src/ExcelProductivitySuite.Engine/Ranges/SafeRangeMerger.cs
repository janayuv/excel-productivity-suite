using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum SafeMergeDirection
    {
        RowByRow,          // Merge across columns for each row
        ColumnByColumn,    // Merge down rows for each column
        EntireSelection    // Merge all cells into top-left cell
    }

    public class SafeMergeOptions
    {
        public SafeMergeDirection Direction { get; set; } = SafeMergeDirection.RowByRow;
        public string Separator { get; set; } = " ";
        public bool SkipBlankCells { get; set; } = true;
        public bool TrimValues { get; set; } = true;
    }

    /// <summary>
    /// Prepares cell ranges for merging without losing data.
    /// Combines values into the anchor cell and clears sibling cells.
    /// </summary>
    public static class SafeRangeMerger
    {
        public static object?[,] PrepareMerge(object?[,] source, SafeMergeOptions options, out int mergedGroups)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            object?[,] result = new object?[rowCount, colCount];
            mergedGroups = 0;

            string sep = options.Separator ?? "";

            switch (options.Direction)
            {
                case SafeMergeDirection.RowByRow:
                    for (int r = 0; r < rowCount; r++)
                    {
                        var tokens = new List<string>();
                        for (int c = 0; c < colCount; c++)
                        {
                            object? val = source[rStart + r, cStart + c];
                            string s = val?.ToString() ?? "";
                            if (options.TrimValues) s = s.Trim();

                            if (!string.IsNullOrEmpty(s) || !options.SkipBlankCells)
                            {
                                tokens.Add(s);
                            }
                        }

                        result[r, 0] = tokens.Count > 0 ? string.Join(sep, tokens) : null;
                        for (int c = 1; c < colCount; c++)
                        {
                            result[r, c] = null; // Cleared for merge
                        }
                        mergedGroups++;
                    }
                    break;

                case SafeMergeDirection.ColumnByColumn:
                    for (int c = 0; c < colCount; c++)
                    {
                        var tokens = new List<string>();
                        for (int r = 0; r < rowCount; r++)
                        {
                            object? val = source[rStart + r, cStart + c];
                            string s = val?.ToString() ?? "";
                            if (options.TrimValues) s = s.Trim();

                            if (!string.IsNullOrEmpty(s) || !options.SkipBlankCells)
                            {
                                tokens.Add(s);
                            }
                        }

                        result[0, c] = tokens.Count > 0 ? string.Join(sep, tokens) : null;
                        for (int r = 1; r < rowCount; r++)
                        {
                            result[r, c] = null; // Cleared for merge
                        }
                        mergedGroups++;
                    }
                    break;

                case SafeMergeDirection.EntireSelection:
                    var allTokens = new List<string>();
                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            object? val = source[rStart + r, cStart + c];
                            string s = val?.ToString() ?? "";
                            if (options.TrimValues) s = s.Trim();

                            if (!string.IsNullOrEmpty(s) || !options.SkipBlankCells)
                            {
                                allTokens.Add(s);
                            }
                        }
                    }

                    result[0, 0] = allTokens.Count > 0 ? string.Join(sep, allTokens) : null;
                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = (r == 0 ? 1 : 0); c < colCount; c++)
                        {
                            result[r, c] = null;
                        }
                    }
                    mergedGroups = 1;
                    break;
            }

            return result;
        }
    }
}
