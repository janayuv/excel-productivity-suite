using System;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum ExtractionMode
    {
        BeforeDelimiter,
        AfterDelimiter,
        BetweenDelimiters
    }

    public enum DelimiterOccurrence
    {
        First,
        Last,
        Nth
    }

    public class TextExtractionOptions
    {
        public ExtractionMode Mode { get; set; } = ExtractionMode.BeforeDelimiter;
        public string PrimaryDelimiter { get; set; } = ",";
        public string SecondaryDelimiter { get; set; } = "";
        public DelimiterOccurrence Occurrence { get; set; } = DelimiterOccurrence.First;
        public int NthIndex { get; set; } = 1;
        public bool CaseSensitive { get; set; } = false;
        public bool ReturnEmptyIfNotFound { get; set; } = true;
    }

    public static class TextExtractor
    {
        public static string Extract(string input, TextExtractionOptions options)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            if (options == null) throw new ArgumentNullException(nameof(options));

            var comparison = options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            switch (options.Mode)
            {
                case ExtractionMode.BeforeDelimiter:
                    return ExtractBefore(input, options.PrimaryDelimiter, options.Occurrence, options.NthIndex, comparison, options.ReturnEmptyIfNotFound);

                case ExtractionMode.AfterDelimiter:
                    return ExtractAfter(input, options.PrimaryDelimiter, options.Occurrence, options.NthIndex, comparison, options.ReturnEmptyIfNotFound);

                case ExtractionMode.BetweenDelimiters:
                    return ExtractBetween(input, options.PrimaryDelimiter, options.SecondaryDelimiter, comparison, options.ReturnEmptyIfNotFound);

                default:
                    return input;
            }
        }

        private static string ExtractBefore(string text, string delimiter, DelimiterOccurrence occurrence, int nth, StringComparison comparison, bool returnEmpty)
        {
            if (string.IsNullOrEmpty(delimiter)) return text;

            int index = FindDelimiterIndex(text, delimiter, occurrence, nth, comparison);
            if (index < 0) return returnEmpty ? string.Empty : text;

            return text.Substring(0, index);
        }

        private static string ExtractAfter(string text, string delimiter, DelimiterOccurrence occurrence, int nth, StringComparison comparison, bool returnEmpty)
        {
            if (string.IsNullOrEmpty(delimiter)) return text;

            int index = FindDelimiterIndex(text, delimiter, occurrence, nth, comparison);
            if (index < 0) return returnEmpty ? string.Empty : text;

            return text.Substring(index + delimiter.Length);
        }

        private static string ExtractBetween(string text, string startDelim, string endDelim, StringComparison comparison, bool returnEmpty)
        {
            if (string.IsNullOrEmpty(startDelim) || string.IsNullOrEmpty(endDelim)) return text;

            int startIndex = text.IndexOf(startDelim, comparison);
            if (startIndex < 0) return returnEmpty ? string.Empty : text;

            int contentStart = startIndex + startDelim.Length;
            int endIndex = text.IndexOf(endDelim, contentStart, comparison);
            if (endIndex < 0) return returnEmpty ? string.Empty : text;

            return text.Substring(contentStart, endIndex - contentStart);
        }

        private static int FindDelimiterIndex(string text, string delimiter, DelimiterOccurrence occurrence, int nth, StringComparison comparison)
        {
            if (occurrence == DelimiterOccurrence.Last)
            {
                return text.LastIndexOf(delimiter, comparison);
            }

            int count = 0;
            int searchStart = 0;
            int targetOccurrence = occurrence == DelimiterOccurrence.Nth ? Math.Max(1, nth) : 1;

            while (searchStart < text.Length)
            {
                int found = text.IndexOf(delimiter, searchStart, comparison);
                if (found < 0) break;

                count++;
                if (count == targetOccurrence) return found;

                searchStart = found + Math.Max(1, delimiter.Length);
            }

            return -1;
        }
    }
}
