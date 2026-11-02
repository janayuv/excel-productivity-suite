using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelProductivitySuite.Engine.DataValidation
{
    public static class DropdownListBuilder
    {
        public const int MaxInlineListLength = 255;

        public static string BuildListString(string rawInput, bool deduplicate = true)
        {
            if (string.IsNullOrWhiteSpace(rawInput)) return string.Empty;

            // Split on comma or newline
            var separators = new[] { ',', '\r', '\n' };
            var rawItems = rawInput.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            var items = new List<string>(rawItems.Length);
            var seen = deduplicate ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : null;

            foreach (var item in rawItems)
            {
                string trimmed = item.Trim();
                if (string.IsNullOrEmpty(trimmed)) continue;

                if (seen != null)
                {
                    if (seen.Add(trimmed))
                    {
                        items.Add(trimmed);
                    }
                }
                else
                {
                    items.Add(trimmed);
                }
            }

            return string.Join(",", items);
        }

        public static bool ValidateList(string formattedList, bool isFormula, out string? errorMessage)
        {
            errorMessage = null;

            if (string.IsNullOrWhiteSpace(formattedList))
            {
                errorMessage = "Dropdown list items cannot be empty.";
                return false;
            }

            if (!isFormula && formattedList.Length > MaxInlineListLength)
            {
                errorMessage = $"The total character length of the list ({formattedList.Length} characters) exceeds Excel's limit of {MaxInlineListLength} characters for inline lists. Please shorten items or use a range reference.";
                return false;
            }

            return true;
        }
    }
}
