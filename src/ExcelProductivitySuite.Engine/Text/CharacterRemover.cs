using System;
using System.Text;

namespace ExcelProductivitySuite.Engine.Text
{
    public class CharacterRemovalOptions
    {
        public bool RemoveNumeric { get; set; } = false;
        public bool RemoveAlphabetic { get; set; } = false;
        public bool RemoveNonAlphanumeric { get; set; } = false;
        public bool RemoveNonPrintable { get; set; } = false;
        public bool RemoveCustomCharacters { get; set; } = false;
        public string CustomCharacters { get; set; } = "";
    }

    public static class CharacterRemover
    {
        public static string RemoveCharacters(string input, CharacterRemovalOptions options)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            if (options == null) throw new ArgumentNullException(nameof(options));

            var sb = new StringBuilder(input.Length);
            foreach (char c in input)
            {
                if (options.RemoveNumeric && char.IsDigit(c))
                {
                    continue;
                }

                if (options.RemoveAlphabetic && char.IsLetter(c))
                {
                    continue;
                }

                if (options.RemoveNonAlphanumeric && !char.IsLetterOrDigit(c) && !char.IsWhiteSpace(c))
                {
                    continue;
                }

                if (options.RemoveNonPrintable && (char.IsControl(c) || c == 0x7F))
                {
                    continue;
                }

                if (options.RemoveCustomCharacters && !string.IsNullOrEmpty(options.CustomCharacters) && options.CustomCharacters.IndexOf(c) >= 0)
                {
                    continue;
                }

                sb.Append(c);
            }

            return sb.ToString();
        }
    }
}
