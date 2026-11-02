using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum NumberExtractionMode
    {
        FirstOnly,
        AllJoined,
        Sum
    }

    public class NumberExtractionOptions
    {
        public NumberExtractionMode Mode { get; set; } = NumberExtractionMode.AllJoined;
        public bool AllowDecimals { get; set; } = true;
        public bool AllowNegative { get; set; } = true;
        public string Delimiter { get; set; } = ", ";
    }

    public static class NumberExtractor
    {
        public static string ExtractFromText(string input, NumberExtractionOptions options)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            if (options == null) options = new NumberExtractionOptions();

            string pattern;
            if (options.AllowDecimals && options.AllowNegative)
            {
                pattern = @"-?\d+(?:\.\d+)?";
            }
            else if (options.AllowDecimals)
            {
                pattern = @"\d+(?:\.\d+)?";
            }
            else if (options.AllowNegative)
            {
                pattern = @"-?\d+";
            }
            else
            {
                pattern = @"\d+";
            }

            var matches = Regex.Matches(input, pattern);
            if (matches.Count == 0) return string.Empty;

            switch (options.Mode)
            {
                case NumberExtractionMode.FirstOnly:
                    return matches[0].Value;

                case NumberExtractionMode.Sum:
                    double sum = 0;
                    foreach (Match m in matches)
                    {
                        if (double.TryParse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                        {
                            sum += val;
                        }
                    }
                    return sum.ToString(CultureInfo.InvariantCulture);

                case NumberExtractionMode.AllJoined:
                default:
                    var list = new List<string>(matches.Count);
                    foreach (Match m in matches)
                    {
                        list.Add(m.Value);
                    }
                    return string.Join(options.Delimiter ?? " ", list);
            }
        }

        public static object?[,] ExtractFromMatrix(object?[,] matrix, NumberExtractionOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rows = rEnd - rStart + 1;
            int cols = cEnd - cStart + 1;
            var result = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    object? cellVal = matrix[rStart + r, cStart + c];
                    if (cellVal == null)
                    {
                        result[r, c] = null;
                        continue;
                    }

                    string text = cellVal.ToString() ?? "";
                    string extracted = ExtractFromText(text, options);
                    result[r, c] = string.IsNullOrEmpty(extracted) ? (object?)null : extracted;
                }
            }

            return result;
        }
    }
}
