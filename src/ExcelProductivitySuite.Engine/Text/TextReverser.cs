using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum TextReverseMode
    {
        ByCharacter,
        ByWord,
        ByDelimiter
    }

    public class TextReverseOptions
    {
        public TextReverseMode Mode { get; set; } = TextReverseMode.ByCharacter;
        public string Delimiter { get; set; } = ",";
        public bool SkipBlankCells { get; set; } = true;
    }

    public static class TextReverser
    {
        public static string Reverse(string input, TextReverseOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }

            switch (options.Mode)
            {
                case TextReverseMode.ByCharacter:
                    return ReverseByCharacter(input);

                case TextReverseMode.ByWord:
                    return ReverseByWord(input);

                case TextReverseMode.ByDelimiter:
                    return ReverseByDelimiter(input, options.Delimiter);

                default:
                    return ReverseByCharacter(input);
            }
        }

        private static string ReverseByCharacter(string input)
        {
            // Use StringInfo to safely handle Unicode surrogate pairs and combining characters
            var enumerator = StringInfo.GetTextElementEnumerator(input);
            var elements = new List<string>();

            while (enumerator.MoveNext())
            {
                elements.Add(enumerator.GetTextElement());
            }

            var sb = new StringBuilder(input.Length);
            for (int i = elements.Count - 1; i >= 0; i--)
            {
                sb.Append(elements[i]);
            }

            return sb.ToString();
        }

        private static string ReverseByWord(string input)
        {
            string[] words = input.Split(new[] { ' ' }, StringSplitOptions.None);
            Array.Reverse(words);
            return string.Join(" ", words);
        }

        private static string ReverseByDelimiter(string input, string delimiter)
        {
            if (string.IsNullOrEmpty(delimiter))
            {
                return ReverseByCharacter(input);
            }

            string[] tokens = input.Split(new[] { delimiter }, StringSplitOptions.None);
            Array.Reverse(tokens);
            return string.Join(delimiter, tokens);
        }

        public static object?[,] ReverseRange(object?[,] matrix, TextReverseOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            var result = (object?[,])matrix.Clone();

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    var val = matrix[r, c];
                    if (val != null)
                    {
                        string str = val.ToString() ?? "";
                        if (!string.IsNullOrEmpty(str))
                        {
                            result[r, c] = Reverse(str, options);
                        }
                    }
                }
            }

            return result;
        }
    }
}
