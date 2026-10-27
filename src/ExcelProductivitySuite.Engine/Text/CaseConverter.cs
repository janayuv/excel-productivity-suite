using System;
using System.Globalization;
using System.Text;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum TextCaseMode
    {
        UpperCase,
        LowerCase,
        ProperCase,
        SentenceCase,
        ToggleCase
    }

    public static class CaseConverter
    {
        public static string ConvertCase(string input, TextCaseMode mode)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            switch (mode)
            {
                case TextCaseMode.UpperCase:
                    return input.ToUpper(CultureInfo.CurrentCulture);

                case TextCaseMode.LowerCase:
                    return input.ToLower(CultureInfo.CurrentCulture);

                case TextCaseMode.ProperCase:
                    return ToProperCase(input);

                case TextCaseMode.SentenceCase:
                    return ToSentenceCase(input);

                case TextCaseMode.ToggleCase:
                    return ToToggleCase(input);

                default:
                    return input;
            }
        }

        private static string ToProperCase(string text)
        {
            var sb = new StringBuilder(text.Length);
            bool newWord = true;

            foreach (char c in text)
            {
                if (char.IsWhiteSpace(c) || char.IsPunctuation(c))
                {
                    sb.Append(c);
                    newWord = true;
                }
                else if (newWord)
                {
                    sb.Append(char.ToUpper(c, CultureInfo.CurrentCulture));
                    newWord = false;
                }
                else
                {
                    sb.Append(char.ToLower(c, CultureInfo.CurrentCulture));
                }
            }

            return sb.ToString();
        }

        private static string ToSentenceCase(string text)
        {
            var sb = new StringBuilder(text.Length);
            bool capitalizeNext = true;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (char.IsLetter(c))
                {
                    if (capitalizeNext)
                    {
                        sb.Append(char.ToUpper(c, CultureInfo.CurrentCulture));
                        capitalizeNext = false;
                    }
                    else
                    {
                        sb.Append(char.ToLower(c, CultureInfo.CurrentCulture));
                    }
                }
                else
                {
                    sb.Append(c);
                    if (c == '.' || c == '!' || c == '?')
                    {
                        capitalizeNext = true;
                    }
                }
            }

            return sb.ToString();
        }

        private static string ToToggleCase(string text)
        {
            var sb = new StringBuilder(text.Length);
            foreach (char c in text)
            {
                if (char.IsUpper(c))
                {
                    sb.Append(char.ToLower(c, CultureInfo.CurrentCulture));
                }
                else if (char.IsLower(c))
                {
                    sb.Append(char.ToUpper(c, CultureInfo.CurrentCulture));
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }
    }
}
