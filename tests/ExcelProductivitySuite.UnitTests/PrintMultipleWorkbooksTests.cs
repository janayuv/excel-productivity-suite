using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Printing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class PrintMultipleWorkbooksTests
    {
        [TestMethod]
        public void PrintWorkbooksEngine_ResolveAllowedExtensions_ResolvesCorrectSets()
        {
            var excelOnly = PrintWorkbooksEngine.ResolveAllowedExtensions(true, false, false, null);
            Assert.IsTrue(excelOnly.Contains(".xlsx"));
            Assert.IsTrue(excelOnly.Contains(".xls"));
            Assert.IsTrue(excelOnly.Contains(".xlsm"));
            Assert.IsFalse(excelOnly.Contains(".csv"));

            var all = PrintWorkbooksEngine.ResolveAllowedExtensions(true, true, true, "*.log; .bak");
            Assert.IsTrue(all.Contains(".xlsx"));
            Assert.IsTrue(all.Contains(".csv"));
            Assert.IsTrue(all.Contains(".txt"));
            Assert.IsTrue(all.Contains(".log"));
            Assert.IsTrue(all.Contains(".bak"));
        }

        [TestMethod]
        public void PrintWorkbooksEngine_ScanFiles_FiltersByExtensionAndIgnoresTempFiles()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "EPS_PrintTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);

            try
            {
                File.WriteAllText(Path.Combine(tempDir, "Workbook1.xlsx"), "test content");
                File.WriteAllText(Path.Combine(tempDir, "Workbook2.xls"), "test content");
                File.WriteAllText(Path.Combine(tempDir, "~$Workbook1.xlsx"), "lock file");
                File.WriteAllText(Path.Combine(tempDir, "Notes.txt"), "some notes");
                File.WriteAllText(Path.Combine(tempDir, "Readme.md"), "markdown");

                var items = PrintWorkbooksEngine.ScanFiles(tempDir, recursive: false, includeExcel: true, includeCsv: false, includeText: false, customFilter: null);

                Assert.AreEqual(2, items.Count);
                Assert.IsTrue(items.Any(i => i.FileName == "Workbook1.xlsx"));
                Assert.IsTrue(items.Any(i => i.FileName == "Workbook2.xls"));
                Assert.IsFalse(items.Any(i => i.FileName.StartsWith("~$")));
                Assert.IsFalse(items.Any(i => i.FileName == "Notes.txt"));
                Assert.IsFalse(items.Any(i => i.FileName == "Readme.md"));

                // Formatted file size
                Assert.IsFalse(string.IsNullOrWhiteSpace(items[0].FileSizeFormatted));
            }
            finally
            {
                if (Directory.Exists(tempDir))
                {
                    try { Directory.Delete(tempDir, true); } catch { }
                }
            }
        }

        [TestMethod]
        public void PrintWorkbooksEngine_ShouldPrintSheet_RespectsAllScopes()
        {
            // Entire workbook
            Assert.IsTrue(PrintWorkbooksEngine.ShouldPrintSheet("Sheet1", isChartSheet: false, PrintSheetScope.EntireWorkbook, null));
            Assert.IsTrue(PrintWorkbooksEngine.ShouldPrintSheet("Chart1", isChartSheet: true, PrintSheetScope.EntireWorkbook, null));

            // Worksheets only
            Assert.IsTrue(PrintWorkbooksEngine.ShouldPrintSheet("Sheet1", isChartSheet: false, PrintSheetScope.WorksheetsOnly, null));
            Assert.IsFalse(PrintWorkbooksEngine.ShouldPrintSheet("Chart1", isChartSheet: true, PrintSheetScope.WorksheetsOnly, null));

            // Chart sheets only
            Assert.IsFalse(PrintWorkbooksEngine.ShouldPrintSheet("Sheet1", isChartSheet: false, PrintSheetScope.ChartSheetsOnly, null));
            Assert.IsTrue(PrintWorkbooksEngine.ShouldPrintSheet("Chart1", isChartSheet: true, PrintSheetScope.ChartSheetsOnly, null));

            // Specific sheets (exact, prefix wildcard, suffix wildcard)
            var targets = new[] { "Summary", "Q4_*", "*Report" };
            Assert.IsTrue(PrintWorkbooksEngine.ShouldPrintSheet("Summary", isChartSheet: false, PrintSheetScope.SpecificSheets, targets));
            Assert.IsTrue(PrintWorkbooksEngine.ShouldPrintSheet("summary", isChartSheet: false, PrintSheetScope.SpecificSheets, targets)); // case-insensitive
            Assert.IsTrue(PrintWorkbooksEngine.ShouldPrintSheet("Q4_Financials", isChartSheet: false, PrintSheetScope.SpecificSheets, targets));
            Assert.IsTrue(PrintWorkbooksEngine.ShouldPrintSheet("AnnualReport", isChartSheet: false, PrintSheetScope.SpecificSheets, targets));
            Assert.IsFalse(PrintWorkbooksEngine.ShouldPrintSheet("Budget2026", isChartSheet: false, PrintSheetScope.SpecificSheets, targets));
        }

        [TestMethod]
        public void PrintWorkbooksEngine_ResolvePageRange_CalculatesAccurately()
        {
            var (allFrom, allTo) = PrintWorkbooksEngine.ResolvePageRange(PrintPageRangeMode.AllPages, 1, 5);
            Assert.IsNull(allFrom);
            Assert.IsNull(allTo);

            var (firstFrom, firstTo) = PrintWorkbooksEngine.ResolvePageRange(PrintPageRangeMode.FirstPageOnly, 1, 5);
            Assert.AreEqual(1, firstFrom);
            Assert.AreEqual(1, firstTo);

            var (customFrom, customTo) = PrintWorkbooksEngine.ResolvePageRange(PrintPageRangeMode.CustomRange, 2, 7);
            Assert.AreEqual(2, customFrom);
            Assert.AreEqual(7, customTo);
        }

        [TestMethod]
        public void PrintWorkbooksEngine_ValidateConfig_ValidatesCorrectly()
        {
            // Null config
            var (v1, e1) = PrintWorkbooksEngine.ValidateConfig(null!);
            Assert.IsFalse(v1);
            Assert.IsNotNull(e1);

            // Empty files
            var config = new PrintWorkbooksConfig();
            var (v2, e2) = PrintWorkbooksEngine.ValidateConfig(config);
            Assert.IsFalse(v2);
            Assert.IsTrue(e2!.Contains("at least one workbook"));

            // Non-existent file
            config.SelectedFiles.Add(new WorkbookPrintItem
            {
                FilePath = "C:\\DoesNotExist_123456789.xlsx",
                FileName = "DoesNotExist_123456789.xlsx",
                IsSelected = true
            });
            var (v3, e3) = PrintWorkbooksEngine.ValidateConfig(config);
            Assert.IsFalse(v3);
            Assert.IsTrue(e3!.Contains("File does not exist"));

            // Valid temp file
            string tempFile = Path.GetTempFileName();
            try
            {
                config.SelectedFiles.Clear();
                config.SelectedFiles.Add(new WorkbookPrintItem
                {
                    FilePath = tempFile,
                    FileName = Path.GetFileName(tempFile),
                    IsSelected = true
                });

                var (v4, _) = PrintWorkbooksEngine.ValidateConfig(config);
                Assert.IsTrue(v4);

                // Invalid specific sheets
                config.SheetScope = PrintSheetScope.SpecificSheets;
                config.SpecificSheetNames.Clear();
                var (v5, _) = PrintWorkbooksEngine.ValidateConfig(config);
                Assert.IsFalse(v5);

                config.SpecificSheetNames.Add("Sheet1");
                var (v6, _) = PrintWorkbooksEngine.ValidateConfig(config);
                Assert.IsTrue(v6);

                // Invalid copies
                config.Copies = 150;
                var (v7, _) = PrintWorkbooksEngine.ValidateConfig(config);
                Assert.IsFalse(v7);
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    try { File.Delete(tempFile); } catch { }
                }
            }
        }

        [TestMethod]
        public void PrintWorkbooksEngine_GenerateSummaryMatrix_GeneratesValidGrid()
        {
            var res = new PrintWorkbooksResult
            {
                TotalWorkbooks = 2,
                PrintedWorkbooks = 2,
                TotalSheetsPrinted = 5,
                Duration = TimeSpan.FromSeconds(3.5)
            };

            var items = new List<WorkbookPrintItem>
            {
                new WorkbookPrintItem
                {
                    FileName = "Sales.xlsx",
                    Directory = "C:\\Reports",
                    FileSizeFormatted = "120 KB",
                    SheetsPrintedCount = 3,
                    Status = PrintItemStatus.Completed
                },
                new WorkbookPrintItem
                {
                    FileName = "Expenses.xlsx",
                    Directory = "C:\\Reports",
                    FileSizeFormatted = "85 KB",
                    SheetsPrintedCount = 2,
                    Status = PrintItemStatus.Completed
                }
            };

            var matrix = PrintWorkbooksEngine.GenerateSummaryMatrix(res, items);
            Assert.IsNotNull(matrix);
            Assert.AreEqual(items.Count + 4, matrix.GetLength(0));
            Assert.AreEqual(7, matrix.GetLength(1));

            // Header check
            Assert.AreEqual("Workbook Name", matrix[0, 1]);
            Assert.AreEqual("Sales.xlsx", matrix[1, 1]);
            Assert.AreEqual(3, matrix[1, 4]);
        }

        [TestMethod]
        public void MockExcelContext_BatchPrintWorkbooks_ExecutesAndReportsProgress()
        {
            var ctx = new MockExcelContext();

            // Check printers
            var printers = ctx.GetInstalledPrinterNames();
            Assert.IsTrue(printers.Count > 0);
            Assert.IsNotNull(ctx.GetDefaultPrinterName());

            string tempFile1 = Path.GetTempFileName();
            string tempFile2 = Path.GetTempFileName();

            try
            {
                var config = new PrintWorkbooksConfig
                {
                    SelectedFiles = new List<WorkbookPrintItem>
                    {
                        new WorkbookPrintItem { FilePath = tempFile1, FileName = Path.GetFileName(tempFile1), IsSelected = true },
                        new WorkbookPrintItem { FilePath = tempFile2, FileName = Path.GetFileName(tempFile2), IsSelected = true }
                    }
                };

                int progressEvents = 0;
                var progress = new Progress<(int current, int total, string currentFile)>(_ =>
                {
                    progressEvents++;
                });

                var result = ctx.BatchPrintWorkbooks(config, progress);

                Assert.IsTrue(result.IsSuccess);
                Assert.AreEqual(2, result.TotalWorkbooks);
                Assert.AreEqual(2, result.PrintedWorkbooks);
                Assert.AreEqual(4, result.TotalSheetsPrinted);
                Assert.AreEqual(0, result.FailedWorkbooks);
            }
            finally
            {
                if (File.Exists(tempFile1)) try { File.Delete(tempFile1); } catch { }
                if (File.Exists(tempFile2)) try { File.Delete(tempFile2); } catch { }
            }
        }
    }
}
