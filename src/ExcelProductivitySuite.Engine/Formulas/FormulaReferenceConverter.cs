using System;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Formulas
{
    public enum ReferenceConversionMode
    {
        ToAbsolute,       // $A$1
        ToRowAbsolute,    // A$1
        ToColumnAbsolute, // $A1
        ToRelative        // A1
    }

    /// <summary>
    /// Engine for converting formula cell references between Absolute, Row-Absolute, Column-Absolute, and Relative.
    /// Safely preserves quoted string literals, sheet name qualifiers, operators, and functions.
    /// </summary>
    public static class FormulaReferenceConverter
    {
        private static readonly Regex CellRefRegex = new Regex(
            @"(?<![A-Za-z0-9_])(?<sheet>(?:'(?:[^']|'')+'|[A-Za-z0-9_]+)!)?(?<colLock>\$?)(?<col>[A-Za-z]{1,3})(?<rowLock>\$?)(?<row>[1-9][0-9]{0,6})(?![A-Za-z0-9_]|(?:\s*\())",
            RegexOptions.Compiled);

        private static readonly Regex FullColRegex = new Regex(
            @"(?<![A-Za-z0-9_])(?<sheet>(?:'(?:[^']|'')+'|[A-Za-z0-9_]+)!)?(?<c1Lock>\$?)(?<col1>[A-Za-z]{1,3}):(?<c2Lock>\$?)(?<col2>[A-Za-z]{1,3})(?![A-Za-z0-9_]|(?:\s*\())",
            RegexOptions.Compiled);

        private static readonly Regex FullRowRegex = new Regex(
            @"(?<![A-Za-z0-9_])(?<sheet>(?:'(?:[^']|'')+'|[A-Za-z0-9_]+)!)?(?<r1Lock>\$?)(?<row1>[1-9][0-9]{0,6}):(?<r2Lock>\$?)(?<row2>[1-9][0-9]{0,6})(?![A-Za-z0-9_])",
            RegexOptions.Compiled);

        private static readonly Regex StringLiteralRegex = new Regex(
            @"""(?:[^""]|"""")*""",
            RegexOptions.Compiled);

        /// <summary>
        /// Converts all cell references in a single formula string to the requested reference style.
        /// </summary>
        public static string ConvertFormulaReferences(string formula, ReferenceConversionMode mode)
        {
            if (string.IsNullOrWhiteSpace(formula)) return formula;

            // Only process formula expressions
            bool hasEqualPrefix = formula.StartsWith("=");
            string expression = hasEqualPrefix ? formula.Substring(1) : formula;

            // Split into string literals and code segments to avoid mutating strings like "=$A$1"
            var sb = new StringBuilder();
            if (hasEqualPrefix) sb.Append('=');

            int lastIndex = 0;
            var stringMatches = StringLiteralRegex.Matches(expression);

            foreach (Match strMatch in stringMatches)
            {
                if (strMatch.Index > lastIndex)
                {
                    string codePart = expression.Substring(lastIndex, strMatch.Index - lastIndex);
                    sb.Append(ConvertCodePart(codePart, mode));
                }
                sb.Append(strMatch.Value);
                lastIndex = strMatch.Index + strMatch.Length;
            }

            if (lastIndex < expression.Length)
            {
                string trailingCode = expression.Substring(lastIndex);
                sb.Append(ConvertCodePart(trailingCode, mode));
            }

            return sb.ToString();
        }

        private static string ConvertCodePart(string code, ReferenceConversionMode mode)
        {
            // First convert full column references like A:B or $A:$B
            string transformed = FullColRegex.Replace(code, match =>
            {
                string col1 = match.Groups["col1"].Value.ToUpperInvariant();
                string col2 = match.Groups["col2"].Value.ToUpperInvariant();
                if (!IsValidColumn(col1) || !IsValidColumn(col2)) return match.Value;

                string sheet = match.Groups["sheet"].Value;
                switch (mode)
                {
                    case ReferenceConversionMode.ToAbsolute:
                    case ReferenceConversionMode.ToColumnAbsolute:
                        return $"{sheet}${col1}:${col2.Insert(0, "$")}";
                    case ReferenceConversionMode.ToRelative:
                    case ReferenceConversionMode.ToRowAbsolute:
                        return $"{sheet}{col1}:{col2}";
                    default:
                        return match.Value;
                }
            });

            // Convert full row references like 1:5 or $1:$5
            transformed = FullRowRegex.Replace(transformed, match =>
            {
                string r1 = match.Groups["row1"].Value;
                string r2 = match.Groups["row2"].Value;
                if (!int.TryParse(r1, out int row1) || !int.TryParse(r2, out int row2) || row1 > 1048576 || row2 > 1048576)
                    return match.Value;

                string sheet = match.Groups["sheet"].Value;
                switch (mode)
                {
                    case ReferenceConversionMode.ToAbsolute:
                    case ReferenceConversionMode.ToRowAbsolute:
                        return $"{sheet}${r1}:${r2.Insert(0, "$")}";
                    case ReferenceConversionMode.ToRelative:
                    case ReferenceConversionMode.ToColumnAbsolute:
                        return $"{sheet}{r1}:{r2}";
                    default:
                        return match.Value;
                }
            });

            // Convert standard cell references like A1, $A1, A$1, $A$1
            transformed = CellRefRegex.Replace(transformed, match =>
            {
                string col = match.Groups["col"].Value.ToUpperInvariant();
                string rowStr = match.Groups["row"].Value;
                if (!IsValidColumn(col) || !int.TryParse(rowStr, out int rowNum) || rowNum > 1048576)
                {
                    return match.Value;
                }

                string sheet = match.Groups["sheet"].Value;
                switch (mode)
                {
                    case ReferenceConversionMode.ToAbsolute:
                        return $"{sheet}${col}${rowStr}";
                    case ReferenceConversionMode.ToRowAbsolute:
                        return $"{sheet}{col}${rowStr}";
                    case ReferenceConversionMode.ToColumnAbsolute:
                        return $"{sheet}${col}{rowStr}";
                    case ReferenceConversionMode.ToRelative:
                        return $"{sheet}{col}{rowStr}";
                    default:
                        return match.Value;
                }
            });

            return transformed;
        }

        private static bool IsValidColumn(string col)
        {
            if (string.IsNullOrEmpty(col) || col.Length > 3) return false;
            int colNum = 0;
            foreach (char c in col)
            {
                if (c < 'A' || c > 'Z') return false;
                colNum = colNum * 26 + (c - 'A' + 1);
            }
            return colNum >= 1 && colNum <= 16384; // Excel max column XFD
        }

        /// <summary>
        /// Converts formula references across a 2D matrix of cell formula objects.
        /// Returns the new matrix and the number of formulas modified.
        /// </summary>
        public static (object?[,] ResultMatrix, int ModifiedCount) ConvertMatrix(object?[,] matrix, ReferenceConversionMode mode)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rows = rEnd - rStart + 1;
            int cols = cEnd - cStart + 1;

            object?[,] result = new object?[rows, cols];
            int modifiedCount = 0;

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    int outR = r - rStart;
                    int outC = c - cStart;
                    object? cell = matrix[r, c];

                    if (cell is string str && str.StartsWith("="))
                    {
                        string converted = ConvertFormulaReferences(str, mode);
                        result[outR, outC] = converted;
                        if (!string.Equals(str, converted, StringComparison.Ordinal))
                        {
                            modifiedCount++;
                        }
                    }
                    else
                    {
                        result[outR, outC] = cell;
                    }
                }
            }

            return (result, modifiedCount);
        }
    }
}
