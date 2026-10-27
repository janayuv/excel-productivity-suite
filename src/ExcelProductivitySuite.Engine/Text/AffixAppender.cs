using System;

namespace ExcelProductivitySuite.Engine.Text
{
    public class AffixOptions
    {
        public string Prefix { get; set; } = "";
        public string Suffix { get; set; } = "";
        public bool SkipBlankCells { get; set; } = true;
        public bool SkipIfAlreadyContains { get; set; } = true;
    }

    public static class AffixAppender
    {
        public static string ApplyAffixes(string input, AffixOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            if (string.IsNullOrEmpty(input))
            {
                if (options.SkipBlankCells) return string.Empty;
                return (options.Prefix ?? "") + (options.Suffix ?? "");
            }

            string result = input;

            if (!string.IsNullOrEmpty(options.Prefix))
            {
                bool hasPrefix = options.SkipIfAlreadyContains && result.StartsWith(options.Prefix, StringComparison.OrdinalIgnoreCase);
                if (!hasPrefix)
                {
                    result = options.Prefix + result;
                }
            }

            if (!string.IsNullOrEmpty(options.Suffix))
            {
                bool hasSuffix = options.SkipIfAlreadyContains && result.EndsWith(options.Suffix, StringComparison.OrdinalIgnoreCase);
                if (!hasSuffix)
                {
                    result = result + options.Suffix;
                }
            }

            return result;
        }
    }

    public class PositionInsertOptions
    {
        public string TextToInsert { get; set; } = "";
        public int PositionIndex { get; set; } = 1; // 1-based index
        public bool FromEnd { get; set; } = false;
    }

    public static class PositionInserter
    {
        public static string Insert(string input, PositionInsertOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrEmpty(options.TextToInsert)) return input ?? string.Empty;

            string text = input ?? string.Empty;
            int insertIndex;

            if (options.FromEnd)
            {
                int offsetFromEnd = Math.Max(0, options.PositionIndex);
                insertIndex = Math.Max(0, text.Length - offsetFromEnd);
            }
            else
            {
                int offsetFromStart = Math.Max(0, options.PositionIndex - 1);
                insertIndex = Math.Min(text.Length, offsetFromStart);
            }

            return text.Insert(insertIndex, options.TextToInsert);
        }
    }
}
