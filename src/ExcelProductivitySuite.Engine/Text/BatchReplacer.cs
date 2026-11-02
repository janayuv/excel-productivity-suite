using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Text
{
    public class BatchReplaceRule
    {
        public string Find { get; set; } = string.Empty;
        public string ReplaceWith { get; set; } = string.Empty;
        public bool MatchCase { get; set; } = false;
        public bool MatchEntireCell { get; set; } = false;
        public bool UseRegex { get; set; } = false;

        public BatchReplaceRule() { }

        public BatchReplaceRule(string find, string replaceWith, bool matchCase = false, bool matchEntireCell = false, bool useRegex = false)
        {
            Find = find;
            ReplaceWith = replaceWith;
            MatchCase = matchCase;
            MatchEntireCell = matchEntireCell;
            UseRegex = useRegex;
        }
    }

    public static class BatchReplacer
    {
        public static string ReplaceInText(string input, IReadOnlyList<BatchReplaceRule> rules, out int replacementsCount)
        {
            replacementsCount = 0;
            if (string.IsNullOrEmpty(input) || rules == null || rules.Count == 0) return input;

            string current = input;

            foreach (var rule in rules)
            {
                if (string.IsNullOrEmpty(rule.Find)) continue;

                if (rule.MatchEntireCell)
                {
                    var comparison = rule.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                    if (string.Equals(current, rule.Find, comparison))
                    {
                        current = rule.ReplaceWith;
                        replacementsCount++;
                    }
                }
                else if (rule.UseRegex)
                {
                    try
                    {
                        var regexOptions = RegexOptions.None;
                        if (!rule.MatchCase) regexOptions |= RegexOptions.IgnoreCase;
                        var rx = new Regex(rule.Find, regexOptions);
                        int count = rx.Matches(current).Count;
                        if (count > 0)
                        {
                            current = rx.Replace(current, rule.ReplaceWith);
                            replacementsCount += count;
                        }
                    }
                    catch (ArgumentException)
                    {
                        // Invalid regex, skip rule safely
                    }
                }
                else
                {
                    var comparison = rule.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                    int index = 0;
                    while ((index = current.IndexOf(rule.Find, index, comparison)) >= 0)
                    {
                        current = current.Remove(index, rule.Find.Length).Insert(index, rule.ReplaceWith);
                        index += rule.ReplaceWith.Length;
                        replacementsCount++;
                    }
                }
            }

            return current;
        }

        public static object?[,] ReplaceInMatrix(object?[,] matrix, IReadOnlyList<BatchReplaceRule> rules, out int totalReplacements)
        {
            totalReplacements = 0;
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (rules == null || rules.Count == 0) return (object?[,])matrix.Clone();

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
                    string replaced = ReplaceInText(text, rules, out int cellCount);
                    totalReplacements += cellCount;
                    result[r, c] = replaced;
                }
            }

            return result;
        }
    }
}
