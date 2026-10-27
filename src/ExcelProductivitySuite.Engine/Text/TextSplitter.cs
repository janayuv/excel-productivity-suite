using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelProductivitySuite.Engine.Text
{
    public class TextSplitOptions
    {
        public string Delimiter { get; set; } = ",";
        public bool SplitByNewline { get; set; } = false;
        public bool TrimTokens { get; set; } = true;
        public bool RemoveEmptyTokens { get; set; } = false;
    }

    public static class TextSplitter
    {
        public static string[] Split(string input, TextSplitOptions options)
        {
            if (string.IsNullOrEmpty(input)) return Array.Empty<string>();
            if (options == null) throw new ArgumentNullException(nameof(options));

            string[] tokens;
            if (options.SplitByNewline)
            {
                tokens = input.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            }
            else
            {
                if (string.IsNullOrEmpty(options.Delimiter))
                {
                    // Split into characters
                    tokens = input.Select(c => c.ToString()).ToArray();
                }
                else
                {
                    tokens = input.Split(new[] { options.Delimiter }, StringSplitOptions.None);
                }
            }

            var result = new List<string>(tokens.Length);
            foreach (var raw in tokens)
            {
                string token = options.TrimTokens ? raw.Trim() : raw;
                if (options.RemoveEmptyTokens && string.IsNullOrEmpty(token))
                {
                    continue;
                }
                result.Add(token);
            }

            return result.ToArray();
        }
    }
}
