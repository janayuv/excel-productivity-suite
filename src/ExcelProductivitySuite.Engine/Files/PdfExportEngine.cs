using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Files
{
    /// <summary>
    /// Engine for validating PDF export settings, sanitizing file names, and resolving tokenized export naming patterns.
    /// </summary>
    public static class PdfExportEngine
    {
        private static readonly char[] InvalidChars = Path.GetInvalidFileNameChars();

        /// <summary>
        /// Sanitizes a string for safe use in file system names by stripping or replacing invalid characters.
        /// </summary>
        public static string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return "Sheet";

            string clean = string.Join("_", fileName.Split(InvalidChars, StringSplitOptions.RemoveEmptyEntries)).Trim();
            // Also trim leading/trailing periods or spaces which cause issues in Windows Explorer
            clean = clean.Trim('.', ' ');
            return string.IsNullOrWhiteSpace(clean) ? "Sheet" : clean;
        }

        /// <summary>
        /// Resolves a file naming pattern with tokens such as [SheetName], [WorkbookName], [Index], [Date], and [DateTime].
        /// </summary>
        public static string ResolveFileName(string? pattern, string? workbookName, string? sheetName, int index = 1, DateTime? timestamp = null)
        {
            DateTime now = timestamp ?? DateTime.Now;
            string cleanWb = SanitizeFileName(Path.GetFileNameWithoutExtension(workbookName ?? "Workbook"));
            if (string.IsNullOrWhiteSpace(cleanWb)) cleanWb = "Workbook";

            string cleanSheet = SanitizeFileName(sheetName ?? $"Sheet{index}");
            if (string.IsNullOrWhiteSpace(cleanSheet)) cleanSheet = $"Sheet{index}";

            string p = string.IsNullOrWhiteSpace(pattern) ? "[WorkbookName] - [SheetName]" : pattern!;

            string resolved = p
                .Replace("[SheetName]", cleanSheet)
                .Replace("{SheetName}", cleanSheet)
                .Replace("[WorkbookName]", cleanWb)
                .Replace("{WorkbookName}", cleanWb)
                .Replace("[Index]", index.ToString())
                .Replace("{Index}", index.ToString())
                .Replace("[Date]", now.ToString("yyyy-MM-dd"))
                .Replace("{Date}", now.ToString("yyyy-MM-dd"))
                .Replace("[DateTime]", now.ToString("yyyyMMdd_HHmmss"))
                .Replace("{DateTime}", now.ToString("yyyyMMdd_HHmmss"));

            // Strip invalid characters from resolved name (except allow extension dot)
            resolved = SanitizeFileName(resolved);
            if (string.IsNullOrWhiteSpace(resolved)) resolved = cleanSheet;

            return resolved.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? resolved : $"{resolved}.pdf";
        }

        /// <summary>
        /// Validates settings before attempting to invoke Excel COM export operations.
        /// </summary>
        public static OperationResult ValidateSettings(PdfExportSettings settings, IReadOnlyList<string>? availableSheets = null, bool hasSelection = true)
        {
            if (settings == null)
            {
                return OperationResult.Failed("PDF export settings cannot be null.");
            }

            if (settings.Scope == PdfExportScope.Selection && !hasSelection)
            {
                return OperationResult.Failed("No active cell range is selected to export.");
            }

            if (settings.Scope == PdfExportScope.SelectedSheets)
            {
                if (settings.SelectedSheets == null || settings.SelectedSheets.Count == 0)
                {
                    return OperationResult.Failed("Please select at least one worksheet to export.");
                }

                if (availableSheets != null && availableSheets.Count > 0)
                {
                    var sheetSet = new HashSet<string>(availableSheets, StringComparer.OrdinalIgnoreCase);
                    var missing = settings.SelectedSheets.Where(s => !sheetSet.Contains(s)).ToList();
                    if (missing.Count == settings.SelectedSheets.Count)
                    {
                        return OperationResult.Failed("None of the selected worksheets exist in the active workbook.");
                    }
                }
            }

            if (settings.OutputMode == PdfOutputMode.SingleCombinedFile)
            {
                if (string.IsNullOrWhiteSpace(settings.DestinationPath))
                {
                    return OperationResult.Failed("Please specify a destination PDF file path.");
                }

                try
                {
                    string dir = Path.GetDirectoryName(settings.DestinationPath) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(dir))
                    {
                        return OperationResult.Failed("Destination path must include a valid directory.");
                    }

                    string fileName = Path.GetFileName(settings.DestinationPath);
                    if (string.IsNullOrWhiteSpace(fileName))
                    {
                        return OperationResult.Failed("Destination path must specify a PDF file name.");
                    }
                }
                catch (Exception ex)
                {
                    return OperationResult.Failed($"Invalid destination path: {ex.Message}");
                }
            }
            else // SeparateFilesPerSheet
            {
                if (string.IsNullOrWhiteSpace(settings.OutputDirectory))
                {
                    return OperationResult.Failed("Please specify an output folder for separate PDF files.");
                }

                try
                {
                    if (Path.GetInvalidPathChars().Any(c => settings.OutputDirectory.Contains(c)))
                    {
                        return OperationResult.Failed("The specified output folder path contains invalid characters.");
                    }
                }
                catch (Exception ex)
                {
                    return OperationResult.Failed($"Invalid output folder: {ex.Message}");
                }
            }

            if (settings.ExportPageRange)
            {
                if (settings.PageFrom < 1)
                {
                    return OperationResult.Failed("Start page number must be 1 or greater.");
                }

                if (settings.PageTo < settings.PageFrom)
                {
                    return OperationResult.Failed("End page number cannot be less than start page number.");
                }
            }

            return OperationResult.Succeeded("Settings validated successfully.");
        }
    }
}
