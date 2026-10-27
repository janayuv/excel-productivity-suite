using System;
using System.Text;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum WhitespaceTrimMode
    {
        TrimLeadingAndTrailing,
        CollapseExcessInternalSpaces,
        RemoveAllSpaces
    }

    public class WhitespaceOptions
    {
        public WhitespaceTrimMode Mode { get; set; } = WhitespaceTrimMode.CollapseExcessInternalSpaces;
        public bool ConvertNonBreakingSpaces { get; set; } = true;
    }

    public static class WhitespaceNormalizer
    {
        public static string Normalize(string input, WhitespaceOptions options)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            if (options == null) throw new ArgumentNullException(nameof(options));

            string text = input;
            if (options.ConvertNonBreakingSpaces)
            {
                text = text.Replace('\u00A0', ' ');
            }

            switch (options.Mode)
            {
                case WhitespaceTrimMode.TrimLeadingAndTrailing:
                    return text.Trim();

                case WhitespaceTrimMode.RemoveAllSpaces:
                    var sbAll = new StringBuilder(text.Length);
                    foreach (char c in text)
                    {
                        if (!char.IsWhiteSpace(c))
                        {
                            sbAll.Append(c);
                        }
                    }
                    return sbAll.ToString();

                case WhitespaceTrimMode.CollapseExcessInternalSpaces:
                default:
                    text = text.Trim();
                    var sb = new StringBuilder(text.Length);
                    bool lastWasSpace = false;

                    foreach (char c in text)
                    {
                        if (char.IsWhiteSpace(c))
                        {
                            if (!lastWasSpace)
                            {
                                sb.Append(' ');
                                lastWasSpace = true;
                            }
                        }
                        else
                        {
                            sb.Append(c);
                            lastWasSpace = false;
                        }
                    }
                    return sb.ToString();
            }
        }
    }
}
