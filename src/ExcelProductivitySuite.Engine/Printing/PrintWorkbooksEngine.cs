using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Printing
{
    public static class PrintWorkbooksEngine
    {
        public static readonly string[] ExcelExtensions = { ".xlsx", ".xls", ".xlsm", ".xlsb" };
        public static readonly string[] CsvExtensions = { ".csv" };
        public static readonly string[] TextExtensions = { ".txt", ".prn" };

        public static HashSet<string> ResolveAllowedExtensions(bool includeExcel, bool includeCsv, bool includeText, string? customFilter)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (includeExcel)
            {
                foreach (var ext in ExcelExtensions) set.Add(ext);
            }

            if (includeCsv)
            {
                foreach (var ext in CsvExtensions) set.Add(ext);
            }

            if (includeText)
            {
                foreach (var ext in TextExtensions) set.Add(ext);
            }

            if (!string.IsNullOrWhiteSpace(customFilter))
            {
                var parts = customFilter!.Split(new[] { ';', ',', '|', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {
                    string clean = part.Trim();
                    if (!clean.StartsWith("."))
                    {
                        if (clean.StartsWith("*."))
                        {
                            clean = clean.Substring(1);
                        }
                        else
                        {
                            clean = "." + clean;
                        }
                    }
                    if (clean.Length > 1)
                    {
                        set.Add(clean);
                    }
                }
            }

            // Fallback: If nothing was selected, default to standard Excel files
            if (set.Count == 0)
            {
                foreach (var ext in ExcelExtensions) set.Add(ext);
            }

            return set;
        }

        public static List<WorkbookPrintItem> ScanFiles(
            string directory,
            bool recursive,
            bool includeExcel,
            bool includeCsv,
            bool includeText,
            string? customFilter)
        {
            var result = new List<WorkbookPrintItem>();
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                return result;
            }

            var allowedExts = ResolveAllowedExtensions(includeExcel, includeCsv, includeText, customFilter);
            var searchOption = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            try
            {
                var dirInfo = new DirectoryInfo(directory);
                var files = dirInfo.EnumerateFiles("*.*", searchOption);

                foreach (var file in files)
                {
                    // Skip hidden or temporary Office lock files (e.g. ~$Book1.xlsx)
                    if (file.Name.StartsWith("~$") || (file.Attributes & FileAttributes.Hidden) != 0)
                    {
                        continue;
                    }

                    if (allowedExts.Contains(file.Extension))
                    {
                        result.Add(new WorkbookPrintItem
                        {
                            FilePath = file.FullName,
                            FileName = file.Name,
                            Directory = file.DirectoryName ?? string.Empty,
                            SizeBytes = file.Length,
                            FileSizeFormatted = WorkbookPrintItem.FormatFileSize(file.Length),
                            LastModified = file.LastWriteTime,
                            IsSelected = true,
                            Status = PrintItemStatus.Pending
                        });
                    }
                }
            }
            catch { }

            return result;
        }

        public static bool ShouldPrintSheet(
            string sheetName,
            bool isChartSheet,
            PrintSheetScope scope,
            IEnumerable<string>? specificSheets)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            switch (scope)
            {
                case PrintSheetScope.EntireWorkbook:
                    return true;

                case PrintSheetScope.WorksheetsOnly:
                    return !isChartSheet;

                case PrintSheetScope.ChartSheetsOnly:
                    return isChartSheet;

                case PrintSheetScope.WorksheetsAndCharts:
                    return true;

                case PrintSheetScope.SpecificSheets:
                    if (specificSheets == null) return false;
                    foreach (var target in specificSheets)
                    {
                        if (string.IsNullOrWhiteSpace(target)) continue;
                        string cleanTarget = target.Trim();

                        // Exact match (case-insensitive)
                        if (string.Equals(sheetName.Trim(), cleanTarget, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }

                        // Wildcard prefix/suffix support (e.g. "Report*" or "*Summary")
                        if (cleanTarget.EndsWith("*") && cleanTarget.Length > 1)
                        {
                            string prefix = cleanTarget.Substring(0, cleanTarget.Length - 1);
                            if (sheetName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return true;
                        }
                        if (cleanTarget.StartsWith("*") && cleanTarget.Length > 1)
                        {
                            string suffix = cleanTarget.Substring(1);
                            if (sheetName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) return true;
                        }
                    }
                    return false;

                default:
                    return true;
            }
        }

        public static (int? from, int? to) ResolvePageRange(PrintPageRangeMode mode, int fromPage, int toPage)
        {
            switch (mode)
            {
                case PrintPageRangeMode.FirstPageOnly:
                    return (1, 1);

                case PrintPageRangeMode.CustomRange:
                    int from = Math.Max(1, fromPage);
                    int to = Math.Max(from, toPage);
                    return (from, to);

                case PrintPageRangeMode.AllPages:
                default:
                    return (null, null);
            }
        }

        public static (bool isValid, string? errorMessage) ValidateConfig(PrintWorkbooksConfig config)
        {
            if (config == null) return (false, "Configuration cannot be null.");

            var selected = config.SelectedFiles?.Where(f => f.IsSelected).ToList() ?? new List<WorkbookPrintItem>();
            if (selected.Count == 0)
            {
                return (false, "Please select at least one workbook to print.");
            }

            foreach (var item in selected)
            {
                if (!File.Exists(item.FilePath))
                {
                    return (false, $"File does not exist: {item.FilePath}");
                }
            }

            if (config.SheetScope == PrintSheetScope.SpecificSheets)
            {
                if (config.SpecificSheetNames == null || !config.SpecificSheetNames.Any(s => !string.IsNullOrWhiteSpace(s)))
                {
                    return (false, "Please specify at least one sheet name when 'Specific Sheets' is selected.");
                }
            }

            if (config.PageRangeMode == PrintPageRangeMode.CustomRange)
            {
                if (config.FromPage < 1) return (false, "From page must be at least 1.");
                if (config.ToPage < config.FromPage) return (false, "To page must be greater than or equal to From page.");
            }

            if (config.Copies < 1 || config.Copies > 99)
            {
                return (false, "Number of copies must be between 1 and 99.");
            }

            return (true, null);
        }

        public static object?[,] GenerateSummaryMatrix(PrintWorkbooksResult result, IEnumerable<WorkbookPrintItem> items)
        {
            var itemList = items?.ToList() ?? new List<WorkbookPrintItem>();
            int rowCount = itemList.Count + 4; // Headers + rows + summary footer
            int colCount = 7;

            var matrix = new object?[rowCount, colCount];

            // Header row
            matrix[0, 0] = "#";
            matrix[0, 1] = "Workbook Name";
            matrix[0, 2] = "Folder Directory";
            matrix[0, 3] = "File Size";
            matrix[0, 4] = "Sheets Printed";
            matrix[0, 5] = "Status";
            matrix[0, 6] = "Details / Error Message";

            int r = 1;
            foreach (var item in itemList)
            {
                matrix[r, 0] = r;
                matrix[r, 1] = item.FileName;
                matrix[r, 2] = item.Directory;
                matrix[r, 3] = item.FileSizeFormatted;
                matrix[r, 4] = item.SheetsPrintedCount;
                matrix[r, 5] = item.Status.ToString();
                matrix[r, 6] = item.ErrorMessage ?? "OK";
                r++;
            }

            // Summary row
            matrix[r + 1, 0] = "TOTALS:";
            matrix[r + 1, 1] = $"Processed: {result.TotalWorkbooks}";
            matrix[r + 1, 2] = $"Printed: {result.PrintedWorkbooks}";
            matrix[r + 1, 3] = $"Failed: {result.FailedWorkbooks}";
            matrix[r + 1, 4] = $"Total Sheets: {result.TotalSheetsPrinted}";
            matrix[r + 1, 5] = $"Duration: {result.Duration.TotalSeconds:F1}s";
            matrix[r + 1, 6] = result.IsSuccess ? "All workbooks printed successfully." : $"{result.Errors.Count} error(s) occurred.";

            return matrix;
        }
    }
}
