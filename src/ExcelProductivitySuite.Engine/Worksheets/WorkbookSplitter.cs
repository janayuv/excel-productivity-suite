using System;
using System.IO;

namespace ExcelProductivitySuite.Engine.Worksheets
{
    /// <summary>
    /// Utility for sanitizing and computing destination filenames for workbook splitting operations.
    /// </summary>
    public static class WorkbookSplitter
    {
        public static string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return "Sheet";
            var invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("_", fileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries)).Trim();
        }

        public static string ResolveFileName(string namingPattern, string workbookName, string sheetName, string extension)
        {
            string cleanWb = SanitizeFileName(Path.GetFileNameWithoutExtension(workbookName));
            string cleanSheet = SanitizeFileName(sheetName);

            string pattern = string.IsNullOrWhiteSpace(namingPattern) ? "[SheetName]" : namingPattern;
            string resolved = pattern
                .Replace("[SheetName]", cleanSheet)
                .Replace("[WorkbookName]", cleanWb);

            if (string.IsNullOrWhiteSpace(resolved)) resolved = cleanSheet;

            string cleanExt = (extension ?? "").TrimStart('.');
            return $"{resolved}.{cleanExt}";
        }
    }
}
