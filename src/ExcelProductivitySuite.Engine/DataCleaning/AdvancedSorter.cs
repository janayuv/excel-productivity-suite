using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    public enum AdvancedSortCriterion
    {
        TextLengthAscending,
        TextLengthDescending,
        MonthName,
        WeekdayName,
        StandardAscending,
        StandardDescending
    }

    public class AdvancedSortOptions
    {
        public int SortColumnIndex { get; set; } = 0; // 0-based column index
        public bool HasHeaderRow { get; set; } = true;
        public AdvancedSortCriterion Criterion { get; set; } = AdvancedSortCriterion.TextLengthAscending;
    }

    /// <summary>
    /// Pure computation engine for sorting 2D matrices by text length, month/weekday names, or standard order.
    /// </summary>
    public static class AdvancedSorter
    {
        private static readonly Dictionary<string, int> MonthMap = new(StringComparer.OrdinalIgnoreCase)
        {
            { "jan", 1 }, { "january", 1 },
            { "feb", 2 }, { "february", 2 },
            { "mar", 3 }, { "march", 3 },
            { "apr", 4 }, { "april", 4 },
            { "may", 5 },
            { "jun", 6 }, { "june", 6 },
            { "jul", 7 }, { "july", 7 },
            { "aug", 8 }, { "august", 8 },
            { "sep", 9 }, { "september", 9 }, { "sept", 9 },
            { "oct", 10 }, { "october", 10 },
            { "nov", 11 }, { "november", 11 },
            { "dec", 12 }, { "december", 12 }
        };

        private static readonly Dictionary<string, int> WeekdayMap = new(StringComparer.OrdinalIgnoreCase)
        {
            { "sun", 1 }, { "sunday", 1 },
            { "mon", 2 }, { "monday", 2 },
            { "tue", 3 }, { "tuesday", 3 }, { "tues", 3 },
            { "wed", 4 }, { "wednesday", 4 },
            { "thu", 5 }, { "thursday", 5 }, { "thur", 5 }, { "thurs", 5 },
            { "fri", 6 }, { "friday", 6 },
            { "sat", 7 }, { "saturday", 7 }
        };

        public static object?[,] Sort(object?[,] matrix, AdvancedSortOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int totalRows = rEnd - rStart + 1;
            int totalCols = cEnd - cStart + 1;

            if (totalRows <= 1 || totalCols <= 0) return (object?[,])matrix.Clone();

            int dataStartRow = options.HasHeaderRow ? rStart + 1 : rStart;
            int sortCol = cStart + Math.Min(totalCols - 1, Math.Max(0, options.SortColumnIndex));

            // Extract row indices
            var rowIndices = new List<int>();
            for (int r = dataStartRow; r <= rEnd; r++)
            {
                rowIndices.Add(r);
            }

            rowIndices.Sort((rA, rB) =>
            {
                object? valA = matrix[rA, sortCol];
                object? valB = matrix[rB, sortCol];
                string strA = valA?.ToString()?.Trim() ?? "";
                string strB = valB?.ToString()?.Trim() ?? "";

                switch (options.Criterion)
                {
                    case AdvancedSortCriterion.TextLengthAscending:
                        int lenCmp = strA.Length.CompareTo(strB.Length);
                        return lenCmp != 0 ? lenCmp : string.Compare(strA, strB, StringComparison.OrdinalIgnoreCase);

                    case AdvancedSortCriterion.TextLengthDescending:
                        int lenDescCmp = strB.Length.CompareTo(strA.Length);
                        return lenDescCmp != 0 ? lenDescCmp : string.Compare(strA, strB, StringComparison.OrdinalIgnoreCase);

                    case AdvancedSortCriterion.MonthName:
                        int mA = MonthMap.TryGetValue(strA, out int m1) ? m1 : 99;
                        int mB = MonthMap.TryGetValue(strB, out int m2) ? m2 : 99;
                        int mCmp = mA.CompareTo(mB);
                        return mCmp != 0 ? mCmp : string.Compare(strA, strB, StringComparison.OrdinalIgnoreCase);

                    case AdvancedSortCriterion.WeekdayName:
                        int wA = WeekdayMap.TryGetValue(strA, out int w1) ? w1 : 99;
                        int wB = WeekdayMap.TryGetValue(strB, out int w2) ? w2 : 99;
                        int wCmp = wA.CompareTo(wB);
                        return wCmp != 0 ? wCmp : string.Compare(strA, strB, StringComparison.OrdinalIgnoreCase);

                    case AdvancedSortCriterion.StandardAscending:
                        if (double.TryParse(strA, out double numA) && double.TryParse(strB, out double numB))
                            return numA.CompareTo(numB);
                        return string.Compare(strA, strB, StringComparison.OrdinalIgnoreCase);

                    case AdvancedSortCriterion.StandardDescending:
                        if (double.TryParse(strA, out double numA2) && double.TryParse(strB, out double numB2))
                            return numB2.CompareTo(numA2);
                        return string.Compare(strB, strA, StringComparison.OrdinalIgnoreCase);

                    default:
                        return 0;
                }
            });

            // Construct sorted matrix
            object?[,] result = new object?[totalRows, totalCols];
            int destRow = 0;

            // Copy header if present
            if (options.HasHeaderRow)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    result[0, c - cStart] = matrix[rStart, c];
                }
                destRow = 1;
            }

            // Copy sorted rows
            foreach (int sourceRow in rowIndices)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    result[destRow, c - cStart] = matrix[sourceRow, c];
                }
                destRow++;
            }

            return result;
        }
    }
}
