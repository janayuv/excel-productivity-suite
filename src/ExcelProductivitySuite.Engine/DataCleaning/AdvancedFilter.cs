using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    public enum FilterOperator
    {
        Equals,
        NotEquals,
        Contains,
        DoesNotContain,
        StartsWith,
        EndsWith,
        GreaterThan,
        LessThan,
        GreaterThanOrEqual,
        LessThanOrEqual,
        IsBlank,
        IsNotBlank
    }

    public enum FilterConjunction
    {
        And, // Match all conditions
        Or   // Match any condition
    }

    public class FilterCondition
    {
        public int ColumnIndex { get; set; } // 0-based relative to selection
        public FilterOperator Operator { get; set; } = FilterOperator.Equals;
        public string ComparisonValue { get; set; } = "";
    }

    public class AdvancedFilterOptions
    {
        public FilterConjunction Conjunction { get; set; } = FilterConjunction.And;
        public bool HasHeaders { get; set; } = true;
        public bool MatchCase { get; set; } = false;
        public bool TrimWhitespace { get; set; } = true;
        public List<FilterCondition> Conditions { get; set; } = new List<FilterCondition>();
    }

    public class AdvancedFilterResult
    {
        public List<int> MatchingRowIndices { get; } = new List<int>(); // 0-based data row indices
        public List<object?[]> MatchingRows { get; } = new List<object?[]>();
        public object?[]? HeaderRow { get; set; }
        public int TotalEvaluatedRows { get; set; }
        public int MatchCount => MatchingRowIndices.Count;

        public object?[,] ExportToMatrix()
        {
            int cols = HeaderRow != null ? HeaderRow.Length : (MatchingRows.Count > 0 ? MatchingRows[0].Length : 1);
            int rows = (HeaderRow != null ? 1 : 0) + MatchingRows.Count;

            object?[,] matrix = new object?[Math.Max(1, rows), cols];

            int curRow = 0;
            if (HeaderRow != null)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[0, c] = HeaderRow[c];
                }
                curRow = 1;
            }

            for (int i = 0; i < MatchingRows.Count; i++)
            {
                var row = MatchingRows[i];
                for (int c = 0; c < Math.Min(cols, row.Length); c++)
                {
                    matrix[curRow + i, c] = row[c];
                }
            }

            return matrix;
        }
    }

    public static class AdvancedFilter
    {
        public static AdvancedFilterResult Execute(object?[,] matrix, AdvancedFilterOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var result = new AdvancedFilterResult();

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int totalRows = rEnd - rStart + 1;
            int totalCols = cEnd - cStart + 1;

            if (totalRows == 0 || totalCols == 0) return result;

            int dataStartRow = rStart;
            if (options.HasHeaders)
            {
                result.HeaderRow = new object?[totalCols];
                for (int c = 0; c < totalCols; c++)
                {
                    result.HeaderRow[c] = matrix[rStart, cStart + c];
                }
                dataStartRow++;
            }

            int evaluated = 0;
            var stringComparer = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            for (int r = dataStartRow; r <= rEnd; r++)
            {
                evaluated++;
                bool isMatch;

                if (options.Conditions == null || options.Conditions.Count == 0)
                {
                    isMatch = true;
                }
                else if (options.Conjunction == FilterConjunction.And)
                {
                    isMatch = true;
                    foreach (var cond in options.Conditions)
                    {
                        if (!EvaluateCondition(matrix, r, cStart, cEnd, cond, options, stringComparer))
                        {
                            isMatch = false;
                            break;
                        }
                    }
                }
                else // Or
                {
                    isMatch = false;
                    foreach (var cond in options.Conditions)
                    {
                        if (EvaluateCondition(matrix, r, cStart, cEnd, cond, options, stringComparer))
                        {
                            isMatch = true;
                            break;
                        }
                    }
                }

                if (isMatch)
                {
                    result.MatchingRowIndices.Add(r - rStart);
                    object?[] rowValues = new object?[totalCols];
                    for (int c = 0; c < totalCols; c++)
                    {
                        rowValues[c] = matrix[r, cStart + c];
                    }
                    result.MatchingRows.Add(rowValues);
                }
            }

            result.TotalEvaluatedRows = evaluated;
            return result;
        }

        private static bool EvaluateCondition(object?[,] matrix, int row, int colStart, int colEnd, FilterCondition cond, AdvancedFilterOptions options, StringComparison comparison)
        {
            int targetCol = colStart + cond.ColumnIndex;
            if (targetCol > colEnd || targetCol < colStart) return false;

            object? raw = matrix[row, targetCol];
            string cellStr = raw?.ToString() ?? "";
            if (options.TrimWhitespace) cellStr = cellStr.Trim();

            string compStr = cond.ComparisonValue ?? "";
            if (options.TrimWhitespace) compStr = compStr.Trim();

            switch (cond.Operator)
            {
                case FilterOperator.Equals:
                    return string.Equals(cellStr, compStr, comparison);

                case FilterOperator.NotEquals:
                    return !string.Equals(cellStr, compStr, comparison);

                case FilterOperator.Contains:
                    return cellStr.IndexOf(compStr, comparison) >= 0;

                case FilterOperator.DoesNotContain:
                    return cellStr.IndexOf(compStr, comparison) < 0;

                case FilterOperator.StartsWith:
                    return cellStr.StartsWith(compStr, comparison);

                case FilterOperator.EndsWith:
                    return cellStr.EndsWith(compStr, comparison);

                case FilterOperator.IsBlank:
                    return string.IsNullOrWhiteSpace(cellStr);

                case FilterOperator.IsNotBlank:
                    return !string.IsNullOrWhiteSpace(cellStr);

                case FilterOperator.GreaterThan:
                case FilterOperator.LessThan:
                case FilterOperator.GreaterThanOrEqual:
                case FilterOperator.LessThanOrEqual:
                    if (double.TryParse(cellStr, out double cellNum) && double.TryParse(compStr, out double compNum))
                    {
                        return cond.Operator switch
                        {
                            FilterOperator.GreaterThan => cellNum > compNum,
                            FilterOperator.LessThan => cellNum < compNum,
                            FilterOperator.GreaterThanOrEqual => cellNum >= compNum,
                            FilterOperator.LessThanOrEqual => cellNum <= compNum,
                            _ => false
                        };
                    }
                    // String comparison fallback
                    int cmp = string.Compare(cellStr, compStr, comparison == StringComparison.Ordinal ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
                    return cond.Operator switch
                    {
                        FilterOperator.GreaterThan => cmp > 0,
                        FilterOperator.LessThan => cmp < 0,
                        FilterOperator.GreaterThanOrEqual => cmp >= 0,
                        FilterOperator.LessThanOrEqual => cmp <= 0,
                        _ => false
                    };

                default:
                    return false;
            }
        }
    }
}
