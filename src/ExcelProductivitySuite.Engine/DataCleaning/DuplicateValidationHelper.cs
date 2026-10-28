using System;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    /// <summary>
    /// Helper for formatting and validating duplicate prevention Excel formulas.
    /// </summary>
    public static class DuplicateValidationHelper
    {
        public static string BuildCountIfFormula(string columnLetter, int startRow)
        {
            if (string.IsNullOrWhiteSpace(columnLetter)) throw new ArgumentNullException(nameof(columnLetter));
            int row = Math.Max(1, startRow);
            string col = columnLetter.Trim().ToUpperInvariant();
            return $"=COUNTIF({col}:{col}, {col}{row})<=1";
        }

        public static string IndexToColumnLetter(int columnIndex)
        {
            if (columnIndex <= 0) columnIndex = 1;
            int div = columnIndex;
            string colLetter = string.Empty;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return colLetter;
        }
    }
}
