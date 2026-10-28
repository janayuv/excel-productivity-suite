using System;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Text
{
    public class TextCleanOptions
    {
        public bool RemoveControlCodes { get; set; } = true;
        public bool RemoveInvisibleZeroWidth { get; set; } = true;
        public bool DecodeHtmlEntities { get; set; } = true;
        public bool PreserveLineBreaks { get; set; } = false;
        public bool PreserveTabs { get; set; } = false;
    }

    public static class TextCleaner
    {
        // Invisible Unicode characters
        private static readonly char[] ZeroWidthChars = new[]
        {
            '\u200B', // Zero width space
            '\u200C', // Zero width non-joiner
            '\u200D', // Zero width joiner
            '\u200E', // Left-to-right mark
            '\u200F', // Right-to-left mark
            '\uFEFF', // Zero width no-break space (BOM)
            '\u00AD'  // Soft hyphen
        };

        public static string Clean(string input, TextCleanOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrEmpty(input)) return string.Empty;

            string text = input;

            // 1. Decode HTML entities if requested
            if (options.DecodeHtmlEntities && text.IndexOf('&') >= 0)
            {
                text = WebUtility.HtmlDecode(text);
            }

            var sb = new StringBuilder(text.Length);

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                // Check invisible zero-width chars
                if (options.RemoveInvisibleZeroWidth && Array.IndexOf(ZeroWidthChars, c) >= 0)
                {
                    continue;
                }

                // Check ASCII control codes
                if (options.RemoveControlCodes)
                {
                    if (c == '\t' && options.PreserveTabs)
                    {
                        sb.Append(c);
                        continue;
                    }

                    if ((c == '\r' || c == '\n') && options.PreserveLineBreaks)
                    {
                        sb.Append(c);
                        continue;
                    }

                    if (char.IsControl(c) || c == 0x7F)
                    {
                        // Replace line breaks and tabs with a space if not preserving them
                        if ((c == '\r' || c == '\n' || c == '\t'))
                        {
                            if (sb.Length > 0 && sb[sb.Length - 1] != ' ')
                            {
                                sb.Append(' ');
                            }
                        }
                        continue;
                    }
                }

                sb.Append(c);
            }

            return sb.ToString();
        }
    }
}
