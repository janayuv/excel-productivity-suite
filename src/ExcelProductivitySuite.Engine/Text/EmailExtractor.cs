using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum EmailExtractionMode
    {
        FirstOnly,
        AllJoined
    }

    public class EmailExtractionOptions
    {
        public EmailExtractionMode Mode { get; set; } = EmailExtractionMode.FirstOnly;
        public bool Deduplicate { get; set; } = true;
        public string Delimiter { get; set; } = "; ";
    }

    public static class EmailExtractor
    {
        private static readonly Regex EmailRegex = new(
            @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static string ExtractFromText(string input, EmailExtractionOptions options)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            if (options == null) options = new EmailExtractionOptions();

            var matches = EmailRegex.Matches(input);
            if (matches.Count == 0) return string.Empty;

            if (options.Mode == EmailExtractionMode.FirstOnly)
            {
                return matches[0].Value;
            }

            var emails = new List<string>(matches.Count);
            var seen = options.Deduplicate ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : null;

            foreach (Match m in matches)
            {
                string val = m.Value;
                if (seen != null)
                {
                    if (seen.Add(val))
                    {
                        emails.Add(val);
                    }
                }
                else
                {
                    emails.Add(val);
                }
            }

            return string.Join(options.Delimiter ?? "; ", emails);
        }

        public static object?[,] ExtractFromMatrix(object?[,] matrix, EmailExtractionOptions options)
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
