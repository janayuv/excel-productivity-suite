using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelProductivitySuite.Engine.Navigation
{
    /// <summary>
    /// Validates worksheet names according to Microsoft Excel's naming rules.
    /// </summary>
    public static class SheetNameValidator
    {
        private static readonly char[] IllegalChars = { '\\', '/', '?', '*', ':', '[', ']' };

        /// <summary>
        /// Validates a potential worksheet name.
        /// </summary>
        /// <param name="name">Candidate name to validate.</param>
        /// <param name="existingNames">Existing sheet names in the workbook to check for collisions.</param>
        /// <param name="errorMessage">Detailed reason if validation fails.</param>
        /// <returns>True if the name is valid in Excel; otherwise false.</returns>
        public static bool Validate(string? name, IEnumerable<string>? existingNames, out string errorMessage)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                errorMessage = "Worksheet name cannot be blank or whitespace.";
                return false;
            }

            string trimmed = name!.Trim();

            if (trimmed.Length > 31)
            {
                errorMessage = $"Worksheet name cannot exceed 31 characters (current: {trimmed.Length}).";
                return false;
            }

            if (trimmed.IndexOfAny(IllegalChars) >= 0)
            {
                errorMessage = "Worksheet name cannot contain any of the following characters: \\ / ? * : [ ]";
                return false;
            }

            if (trimmed.StartsWith("'", StringComparison.Ordinal) || trimmed.EndsWith("'", StringComparison.Ordinal))
            {
                errorMessage = "Worksheet name cannot begin or end with a single quote (').";
                return false;
            }

            if (string.Equals(trimmed, "History", StringComparison.OrdinalIgnoreCase))
            {
                errorMessage = "'History' is a reserved Excel name and cannot be used.";
                return false;
            }

            if (existingNames != null)
            {
                bool exists = existingNames.Any(existing => string.Equals(existing.Trim(), trimmed, StringComparison.OrdinalIgnoreCase));
                if (exists)
                {
                    errorMessage = $"A worksheet named '{trimmed}' already exists in this workbook.";
                    return false;
                }
            }

            errorMessage = string.Empty;
            return true;
        }

        /// <summary>
        /// Sanitizes an arbitrary string into a valid Excel worksheet name.
        /// </summary>
        public static string Sanitize(string name, string fallback = "Sheet")
        {
            if (string.IsNullOrWhiteSpace(name)) return fallback;

            char[] chars = name.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (IllegalChars.Contains(chars[i]))
                {
                    chars[i] = '_';
                }
            }

            string result = new string(chars).Trim(' ', '\'');
            if (result.Length > 31)
            {
                result = result.Substring(0, 31).TrimEnd();
            }

            if (string.IsNullOrWhiteSpace(result) || string.Equals(result, "History", StringComparison.OrdinalIgnoreCase))
            {
                return fallback;
            }

            return result;
        }
    }
}
