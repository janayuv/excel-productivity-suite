using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum UrlExtractionMode
    {
        FirstOnly,
        AllJoined
    }

    public class UrlExtractionOptions
    {
        public UrlExtractionMode Mode { get; set; } = UrlExtractionMode.FirstOnly;
        public bool Deduplicate { get; set; } = true;
        public string Delimiter { get; set; } = "; ";
        public bool NormalizeScheme { get; set; } = true;
    }

    public static class UrlExtractor
    {
        // Matches http://, https://, ftp://, or www. followed by valid domain and path characters
        private static readonly Regex UrlRegex = new(
            @"(?:https?://|ftp://|www\.)[a-zA-Z0-9\-._~:/?#\[\]@!$&'()*+,;=%]+",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly char[] TrailingPunctuation = { '.', ',', ';', ':', ')', ']', '}', '>', '\'', '"' };

        public static string ExtractFromText(string input, UrlExtractionOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            options ??= new UrlExtractionOptions();

            var matches = UrlRegex.Matches(input);
            if (matches.Count == 0) return string.Empty;

            var cleanedUrls = new List<string>(matches.Count);
            var seen = options.Deduplicate ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : null;

            foreach (Match m in matches)
            {
                string rawUrl = m.Value.TrimEnd(TrailingPunctuation);
                if (string.IsNullOrEmpty(rawUrl)) continue;

                if (options.NormalizeScheme && rawUrl.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
                {
                    rawUrl = "https://" + rawUrl;
                }

                if (seen != null)
                {
                    if (seen.Add(rawUrl))
                    {
                        cleanedUrls.Add(rawUrl);
                    }
                }
                else
                {
                    cleanedUrls.Add(rawUrl);
                }

                if (options.Mode == UrlExtractionMode.FirstOnly && cleanedUrls.Count > 0)
                {
                    return cleanedUrls[0];
                }
            }

            if (cleanedUrls.Count == 0) return string.Empty;

            return options.Mode == UrlExtractionMode.FirstOnly
                ? cleanedUrls[0]
                : string.Join(options.Delimiter ?? "; ", cleanedUrls);
        }

        public static object?[,] ExtractFromMatrix(object?[,] matrix, UrlExtractionOptions? options = null)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            options ??= new UrlExtractionOptions();

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
