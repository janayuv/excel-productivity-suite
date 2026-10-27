using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum SelectionCriteria
    {
        BlankCells,
        NonBlankCells,
        ErrorCells,
        NumericCells,
        TextCells,
        FormulaCells,
        MaxNumericValue,
        MinNumericValue
    }

    public class CellCoordinate
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public object? Value { get; set; }

        public CellCoordinate(int row, int col, object? value)
        {
            Row = row;
            Column = col;
            Value = value;
        }
    }

    public static class RangeSelector
    {
        private static readonly HashSet<string> ExcelErrors = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!", "#SPILL!", "#CALC!"
        };

        public static List<CellCoordinate> FindMatchingCells(object?[,] matrix, int startRow, int startCol, SelectionCriteria criteria)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));

            var results = new List<CellCoordinate>();
            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            double maxVal = double.MinValue;
            double minVal = double.MaxValue;
            var numCells = new List<(int row, int col, double val)>();

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    int sheetRow = startRow + (r - rStart);
                    int sheetCol = startCol + (c - cStart);
                    object? val = matrix[r, c];
                    string str = val?.ToString()?.Trim() ?? "";

                    switch (criteria)
                    {
                        case SelectionCriteria.BlankCells:
                            if (val == null || string.IsNullOrWhiteSpace(str))
                            {
                                results.Add(new CellCoordinate(sheetRow, sheetCol, val));
                            }
                            break;

                        case SelectionCriteria.NonBlankCells:
                            if (val != null && !string.IsNullOrWhiteSpace(str))
                            {
                                results.Add(new CellCoordinate(sheetRow, sheetCol, val));
                            }
                            break;

                        case SelectionCriteria.ErrorCells:
                            if (!string.IsNullOrEmpty(str) && (str.StartsWith("#") || ExcelErrors.Contains(str)))
                            {
                                results.Add(new CellCoordinate(sheetRow, sheetCol, val));
                            }
                            break;

                        case SelectionCriteria.NumericCells:
                            if (double.TryParse(str, out _))
                            {
                                results.Add(new CellCoordinate(sheetRow, sheetCol, val));
                            }
                            break;

                        case SelectionCriteria.TextCells:
                            if (!string.IsNullOrWhiteSpace(str) && !double.TryParse(str, out _) && !str.StartsWith("#"))
                            {
                                results.Add(new CellCoordinate(sheetRow, sheetCol, val));
                            }
                            break;

                        case SelectionCriteria.FormulaCells:
                            if (str.StartsWith("="))
                            {
                                results.Add(new CellCoordinate(sheetRow, sheetCol, val));
                            }
                            break;

                        case SelectionCriteria.MaxNumericValue:
                        case SelectionCriteria.MinNumericValue:
                            if (double.TryParse(str, out double parsedNum))
                            {
                                if (parsedNum > maxVal) maxVal = parsedNum;
                                if (parsedNum < minVal) minVal = parsedNum;
                                numCells.Add((sheetRow, sheetCol, parsedNum));
                            }
                            break;
                    }
                }
            }

            if (criteria == SelectionCriteria.MaxNumericValue && numCells.Count > 0)
            {
                foreach (var item in numCells)
                {
                    if (Math.Abs(item.val - maxVal) < 0.0000000001)
                    {
                        results.Add(new CellCoordinate(item.row, item.col, item.val));
                    }
                }
            }
            else if (criteria == SelectionCriteria.MinNumericValue && numCells.Count > 0)
            {
                foreach (var item in numCells)
                {
                    if (Math.Abs(item.val - minVal) < 0.0000000001)
                    {
                        results.Add(new CellCoordinate(item.row, item.col, item.val));
                    }
                }
            }

            return results;
        }
    }
}
