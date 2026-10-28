using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Worksheets
{
    public class SheetIndexEntry
    {
        public int SequenceNumber { get; set; }
        public string SheetName { get; set; } = "";
        public bool IsVisible { get; set; } = true;
        public string HyperlinkFormula { get; set; } = "";
    }

    /// <summary>
    /// Generates structured table-of-contents entries and Excel HYPERLINK formulas for sheet navigation.
    /// </summary>
    public static class SheetIndexGenerator
    {
        public static List<SheetIndexEntry> GenerateEntries(IEnumerable<string> sheetNames, string indexSheetName)
        {
            if (sheetNames == null) throw new ArgumentNullException(nameof(sheetNames));

            var results = new List<SheetIndexEntry>();
            int seq = 1;

            foreach (var s in sheetNames)
            {
                if (string.Equals(s, indexSheetName, StringComparison.OrdinalIgnoreCase)) continue;

                string safeTarget = s.Replace("'", "''");
                string formula = $"=HYPERLINK(\"#'{safeTarget}'!A1\", \"Go to {safeTarget}\")";

                results.Add(new SheetIndexEntry
                {
                    SequenceNumber = seq++,
                    SheetName = s,
                    IsVisible = true,
                    HyperlinkFormula = formula
                });
            }

            return results;
        }

        public static string CreateBackLinkFormula(string indexSheetName)
        {
            string safeIndex = (indexSheetName ?? "Index").Replace("'", "''");
            return $"=HYPERLINK(\"#'{safeIndex}'!A1\", \"[Back to Index]\")";
        }
    }
}
