using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Worksheets
{
    public enum SheetSortMode
    {
        NaturalNumeric,
        AlphabeticalAscending,
        AlphabeticalDescending,
        ReverseCurrent
    }

    /// <summary>
    /// Pure algorithm engine for ordering worksheet tabs alphabetically, naturally by numeric values, or reversed.
    /// </summary>
    public static class SheetTabSorter
    {
        public static List<string> Sort(IEnumerable<string> sheetNames, SheetSortMode mode)
        {
            if (sheetNames == null) throw new ArgumentNullException(nameof(sheetNames));

            var list = new List<string>(sheetNames);
            if (list.Count <= 1) return list;

            switch (mode)
            {
                case SheetSortMode.AlphabeticalAscending:
                    list.Sort(StringComparer.OrdinalIgnoreCase);
                    break;

                case SheetSortMode.AlphabeticalDescending:
                    list.Sort((a, b) => string.Compare(b, a, StringComparison.OrdinalIgnoreCase));
                    break;

                case SheetSortMode.ReverseCurrent:
                    list.Reverse();
                    break;

                case SheetSortMode.NaturalNumeric:
                    list.Sort(NaturalCompare);
                    break;
            }

            return list;
        }

        /// <summary>
        /// Compares two strings using natural numeric sorting (e.g., "Sheet2" before "Sheet10").
        /// </summary>
        public static int NaturalCompare(string? x, string? y)
        {
            if (ReferenceEquals(x, y)) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            // Split into numeric and non-numeric chunks
            var regex = new Regex(@"(\d+)|(\D+)");
            var xMatches = regex.Matches(x);
            var yMatches = regex.Matches(y);

            int count = Math.Min(xMatches.Count, yMatches.Count);
            for (int i = 0; i < count; i++)
            {
                string xToken = xMatches[i].Value;
                string yToken = yMatches[i].Value;

                bool xIsNum = long.TryParse(xToken, out long xNum);
                bool yIsNum = long.TryParse(yToken, out long yNum);

                int cmp;
                if (xIsNum && yIsNum)
                {
                    cmp = xNum.CompareTo(yNum);
                }
                else
                {
                    cmp = string.Compare(xToken, yToken, StringComparison.OrdinalIgnoreCase);
                }

                if (cmp != 0) return cmp;
            }

            return xMatches.Count.CompareTo(yMatches.Count);
        }
    }
}
