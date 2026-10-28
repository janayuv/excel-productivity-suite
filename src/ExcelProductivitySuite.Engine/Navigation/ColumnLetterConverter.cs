using System;
using System.Text;

namespace ExcelProductivitySuite.Engine.Navigation
{
    public static class ColumnLetterConverter
    {
        public static string IndexToLetter(int columnIndex)
        {
            if (columnIndex < 1) return "A";

            int dividend = columnIndex;
            var sb = new StringBuilder();

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                sb.Insert(0, (char)('A' + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return sb.ToString();
        }

        public static int LetterToIndex(string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter)) return 1;

            string upper = columnLetter.Trim().ToUpperInvariant();
            int sum = 0;

            for (int i = 0; i < upper.Length; i++)
            {
                char c = upper[i];
                if (c >= 'A' && c <= 'Z')
                {
                    sum = (sum * 26) + (c - 'A' + 1);
                }
            }

            return sum > 0 ? sum : 1;
        }
    }
}
