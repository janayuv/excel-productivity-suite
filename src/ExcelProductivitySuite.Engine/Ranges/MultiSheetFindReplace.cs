using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public class FindReplaceOptions
    {
        public string SearchText { get; set; } = "";
        public string ReplaceText { get; set; } = "";
        public bool MatchCase { get; set; } = false;
        public bool MatchEntireCell { get; set; } = false;
        public bool UseRegex { get; set; } = false;
    }

    public class FindMatchResult
    {
        public string SheetName { get; set; } = "";
        public int Row { get; set; }
        public int Column { get; set; }
        public string Address { get; set; } = "";
        public string OriginalText { get; set; } = "";
        public string ReplacedText { get; set; } = "";
    }

    /// <summary>
    /// Engine for scanning, previewing, and replacing cell contents across single or multiple sheets.
    /// </summary>
    public static class MultiSheetFindReplace
    {
        public static List<FindMatchResult> FindMatches(
            string sheetName,
            int startRow,
            int startCol,
            object?[,] matrix,
            FindReplaceOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var results = new List<FindMatchResult>();
            if (string.IsNullOrEmpty(options.SearchText)) return results;

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            Regex? regex = null;
            if (options.UseRegex)
            {
                var regexOpts = options.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                regex = new Regex(options.SearchText, regexOpts);
            }

            var strComparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    object? cellVal = matrix[r, c];
                    if (cellVal == null) continue;

                    string original = cellVal.ToString() ?? "";
                    if (string.IsNullOrEmpty(original)) continue;

                    bool isMatch = false;
                    string replaced = original;

                    if (options.UseRegex && regex != null)
                    {
                        if (options.MatchEntireCell)
                        {
                            var match = regex.Match(original);
                            isMatch = match.Success && match.Length == original.Length;
                        }
                        else
                        {
                            isMatch = regex.IsMatch(original);
                        }

                        if (isMatch)
                        {
                            replaced = regex.Replace(original, options.ReplaceText ?? "");
                        }
                    }
                    else
                    {
                        if (options.MatchEntireCell)
                        {
                            isMatch = string.Equals(original, options.SearchText, strComparison);
                            if (isMatch) replaced = options.ReplaceText ?? "";
                        }
                        else
                        {
                            int idx = original.IndexOf(options.SearchText, strComparison);
                            isMatch = (idx >= 0);
                            if (isMatch)
                            {
                                // Replace all occurrences
                                if (options.MatchCase)
                                {
                                    replaced = original.Replace(options.SearchText, options.ReplaceText ?? "");
                                }
                                else
                                {
                                    replaced = Regex.Replace(
                                        original,
                                        Regex.Escape(options.SearchText),
                                        options.ReplaceText ?? "",
                                        RegexOptions.IgnoreCase);
                                }
                            }
                        }
                    }

                    if (isMatch)
                    {
                        int sheetRow = startRow + (r - rStart);
                        int sheetCol = startCol + (c - cStart);
                        string addr = FormatCellAddress(sheetCol, sheetRow);

                        results.Add(new FindMatchResult
                        {
                            SheetName = sheetName,
                            Row = sheetRow,
                            Column = sheetCol,
                            Address = addr,
                            OriginalText = original,
                            ReplacedText = replaced
                        });
                    }
                }
            }

            return results;
        }

        public static object?[,] ReplaceInMatrix(object?[,] matrix, FindReplaceOptions options, out int replacementCount)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            object?[,] result = (object?[,])matrix.Clone();
            replacementCount = 0;

            if (string.IsNullOrEmpty(options.SearchText)) return result;

            Regex? regex = null;
            if (options.UseRegex)
            {
                var regexOpts = options.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                regex = new Regex(options.SearchText, regexOpts);
            }

            var strComparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    object? cellVal = matrix[r, c];
                    if (cellVal == null) continue;

                    string original = cellVal.ToString() ?? "";
                    if (string.IsNullOrEmpty(original)) continue;

                    bool isMatch = false;
                    string replaced = original;

                    if (options.UseRegex && regex != null)
                    {
                        if (options.MatchEntireCell)
                        {
                            var match = regex.Match(original);
                            isMatch = match.Success && match.Length == original.Length;
                        }
                        else
                        {
                            isMatch = regex.IsMatch(original);
                        }

                        if (isMatch)
                        {
                            replaced = regex.Replace(original, options.ReplaceText ?? "");
                        }
                    }
                    else
                    {
                        if (options.MatchEntireCell)
                        {
                            isMatch = string.Equals(original, options.SearchText, strComparison);
                            if (isMatch) replaced = options.ReplaceText ?? "";
                        }
                        else
                        {
                            int idx = original.IndexOf(options.SearchText, strComparison);
                            isMatch = (idx >= 0);
                            if (isMatch)
                            {
                                if (options.MatchCase)
                                {
                                    replaced = original.Replace(options.SearchText, options.ReplaceText ?? "");
                                }
                                else
                                {
                                    replaced = Regex.Replace(
                                        original,
                                        Regex.Escape(options.SearchText),
                                        options.ReplaceText ?? "",
                                        RegexOptions.IgnoreCase);
                                }
                            }
                        }
                    }

                    if (isMatch && replaced != original)
                    {
                        result[r, c] = replaced;
                        replacementCount++;
                    }
                }
            }

            return result;
        }

        private static string FormatCellAddress(int col, int row)
        {
            string colLetter = "";
            int temp = col;
            while (temp > 0)
            {
                int rem = (temp - 1) % 26;
                colLetter = (char)('A' + rem) + colLetter;
                temp = (temp - 1) / 26;
            }
            return $"${colLetter}${row}";
        }
    }
}
