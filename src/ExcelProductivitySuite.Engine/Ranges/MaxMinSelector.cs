using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum MaxMinTarget
    {
        Maximum,
        Minimum
    }

    public enum MaxMinScope
    {
        EntireSelection,
        EachRow,
        EachColumn
    }

    public class MaxMinSelectionOptions
    {
        public MaxMinTarget Target { get; set; } = MaxMinTarget.Maximum;
        public MaxMinScope Scope { get; set; } = MaxMinScope.EntireSelection;
        public int Count { get; set; } = 1; // Top/Bottom N values
    }

    /// <summary>
    /// Locates the maximum or minimum numeric cells across selections, rows, or columns.
    /// </summary>
    public static class MaxMinSelector
    {
        public static List<(int Row, int Column, double Value)> FindExtremeCells(
            object?[,] matrix,
            int startRow,
            int startCol,
            MaxMinSelectionOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int n = Math.Max(1, options.Count);
            var results = new List<(int, int, double)>();

            switch (options.Scope)
            {
                case MaxMinScope.EntireSelection:
                    var allNumCells = new List<(int row, int col, double val)>();
                    for (int r = rStart; r <= rEnd; r++)
                    {
                        for (int c = cStart; c <= cEnd; c++)
                        {
                            if (TryGetNumeric(matrix[r, c], out double d))
                            {
                                int sheetRow = startRow + (r - rStart);
                                int sheetCol = startCol + (c - cStart);
                                allNumCells.Add((sheetRow, sheetCol, d));
                            }
                        }
                    }

                    if (allNumCells.Count > 0)
                    {
                        var ordered = options.Target == MaxMinTarget.Maximum
                            ? allNumCells.OrderByDescending(x => x.val).Take(n)
                            : allNumCells.OrderBy(x => x.val).Take(n);
                        results.AddRange(ordered);
                    }
                    break;

                case MaxMinScope.EachRow:
                    for (int r = rStart; r <= rEnd; r++)
                    {
                        var rowNumCells = new List<(int row, int col, double val)>();
                        for (int c = cStart; c <= cEnd; c++)
                        {
                            if (TryGetNumeric(matrix[r, c], out double d))
                            {
                                int sheetRow = startRow + (r - rStart);
                                int sheetCol = startCol + (c - cStart);
                                rowNumCells.Add((sheetRow, sheetCol, d));
                            }
                        }

                        if (rowNumCells.Count > 0)
                        {
                            var ordered = options.Target == MaxMinTarget.Maximum
                                ? rowNumCells.OrderByDescending(x => x.val).Take(n)
                                : rowNumCells.OrderBy(x => x.val).Take(n);
                            results.AddRange(ordered);
                        }
                    }
                    break;

                case MaxMinScope.EachColumn:
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        var colNumCells = new List<(int row, int col, double val)>();
                        for (int r = rStart; r <= rEnd; r++)
                        {
                            if (TryGetNumeric(matrix[r, c], out double d))
                            {
                                int sheetRow = startRow + (r - rStart);
                                int sheetCol = startCol + (c - cStart);
                                colNumCells.Add((sheetRow, sheetCol, d));
                            }
                        }

                        if (colNumCells.Count > 0)
                        {
                            var ordered = options.Target == MaxMinTarget.Maximum
                                ? colNumCells.OrderByDescending(x => x.val).Take(n)
                                : colNumCells.OrderBy(x => x.val).Take(n);
                            results.AddRange(ordered);
                        }
                    }
                    break;
            }

            return results;
        }

        private static bool TryGetNumeric(object? value, out double number)
        {
            if (value is double d) { number = d; return true; }
            if (value is int i) { number = i; return true; }
            if (value is long l) { number = l; return true; }
            if (value is decimal m) { number = (double)m; return true; }
            if (value is float f) { number = f; return true; }
            if (value is short s) { number = s; return true; }

            if (value is string str && double.TryParse(str.Trim(), out double parsed))
            {
                number = parsed;
                return true;
            }

            number = 0;
            return false;
        }
    }
}
