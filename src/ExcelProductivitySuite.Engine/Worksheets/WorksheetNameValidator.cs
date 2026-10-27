using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelProductivitySuite.Engine.Worksheets
{
    public class BatchRenameOptions
    {
        public string Prefix { get; set; } = "";
        public string Suffix { get; set; } = "";
        public string FindText { get; set; } = "";
        public string ReplaceWith { get; set; } = "";
        public bool AppendSequenceNumber { get; set; } = false;
        public int SequenceStart { get; set; } = 1;
        public int SequenceStep { get; set; } = 1;
        public string SequenceFormat { get; set; } = "D2";
    }

    public static class WorksheetNameValidator
    {
        private static readonly char[] ForbiddenChars = new[] { '\\', '/', '?', '*', ':', '[', ']' };

        public static (bool IsValid, string? ErrorMessage) ValidateSheetName(string? sheetName, IEnumerable<string>? existingNames = null, string? currentSheetName = null)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return (false, "Worksheet name cannot be empty or only whitespace.");
            }

            string name = sheetName!.Trim();

            if (name.Length > 31)
            {
                return (false, $"Worksheet name exceeds Excel's 31-character limit ({name.Length} chars).");
            }

            if (name.IndexOfAny(ForbiddenChars) >= 0)
            {
                return (false, "Worksheet name contains forbidden characters: \\ / ? * : [ ]");
            }

            if (name.StartsWith("'") || name.EndsWith("'"))
            {
                return (false, "Worksheet name cannot start or end with an apostrophe (').");
            }

            if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase))
            {
                return (false, "'History' is a reserved name in Excel.");
            }

            if (existingNames != null)
            {
                foreach (var existing in existingNames)
                {
                    if (currentSheetName != null && string.Equals(existing, currentSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue; // Renaming sheet to itself is allowed
                    }

                    if (string.Equals(existing, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return (false, $"A worksheet named '{name}' already exists in the workbook.");
                    }
                }
            }

            return (true, null);
        }

        public static string SanitizeSheetName(string proposedName)
        {
            if (string.IsNullOrWhiteSpace(proposedName)) return "Sheet";

            string clean = proposedName.Trim();
            foreach (char c in ForbiddenChars)
            {
                clean = clean.Replace(c, '_');
            }

            clean = clean.Trim('\'', ' ');

            if (clean.Length > 31)
            {
                clean = clean.Substring(0, 31);
            }

            if (string.Equals(clean, "History", StringComparison.OrdinalIgnoreCase))
            {
                clean = "History_1";
            }

            return clean;
        }

        public static Dictionary<string, string> GenerateRenamedNames(IReadOnlyList<string> originalNames, BatchRenameOptions options)
        {
            if (originalNames == null) throw new ArgumentNullException(nameof(originalNames));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int currentSeq = options.SequenceStart;

            for (int i = 0; i < originalNames.Count; i++)
            {
                string oldName = originalNames[i];
                string newName = oldName;

                if (!string.IsNullOrEmpty(options.FindText))
                {
                    newName = newName.Replace(options.FindText, options.ReplaceWith ?? "");
                }

                if (!string.IsNullOrEmpty(options.Prefix))
                {
                    newName = options.Prefix + newName;
                }

                if (!string.IsNullOrEmpty(options.Suffix))
                {
                    newName = newName + options.Suffix;
                }

                if (options.AppendSequenceNumber)
                {
                    string seqStr = currentSeq.ToString(options.SequenceFormat);
                    newName = $"{newName}_{seqStr}";
                    currentSeq += options.SequenceStep;
                }

                newName = SanitizeSheetName(newName);

                // Handle duplicate collision
                string candidate = newName;
                int suffix = 1;
                while (usedNames.Contains(candidate))
                {
                    string num = $"_{suffix++}";
                    if (newName.Length + num.Length > 31)
                    {
                        candidate = newName.Substring(0, 31 - num.Length) + num;
                    }
                    else
                    {
                        candidate = newName + num;
                    }
                }

                usedNames.Add(candidate);
                result[oldName] = candidate;
            }

            return result;
        }
    }
}
