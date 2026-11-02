using System;
using System.Collections.Generic;
using System.IO;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Files;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class PdfExportTests
    {
        [TestMethod]
        public void SanitizeFileName_CleansInvalidCharacters()
        {
            Assert.AreEqual("Sheet", PdfExportEngine.SanitizeFileName(null!));
            Assert.AreEqual("Sheet", PdfExportEngine.SanitizeFileName("   "));
            Assert.AreEqual("Sales_Report", PdfExportEngine.SanitizeFileName("Sales/Report"));
            Assert.AreEqual("Quarter_1_Final", PdfExportEngine.SanitizeFileName("Quarter:1*Final?"));
            Assert.AreEqual("Report", PdfExportEngine.SanitizeFileName("...Report..."));
        }

        [TestMethod]
        public void ResolveFileName_ReplacesAllTokensCorrectly()
        {
            var date = new DateTime(2026, 10, 30, 15, 45, 30);

            string resolved = PdfExportEngine.ResolveFileName(
                "[WorkbookName] - [SheetName] - [Index] - [Date]",
                "AnnualBudget.xlsx",
                "Q1 Summary",
                1,
                date);

            Assert.AreEqual("AnnualBudget - Q1 Summary - 1 - 2026-10-30.pdf", resolved);

            // Test curly brace syntax and datetime
            string resolved2 = PdfExportEngine.ResolveFileName(
                "{WorkbookName}_{SheetName}_{DateTime}",
                "Financials",
                "BalanceSheet",
                2,
                date);

            Assert.AreEqual("Financials_BalanceSheet_20261030_154530.pdf", resolved2);

            // Test default pattern when null/empty
            string defaultResolved = PdfExportEngine.ResolveFileName(
                null,
                "MyBook",
                "Data",
                1,
                date);

            Assert.AreEqual("MyBook - Data.pdf", defaultResolved);
        }

        [TestMethod]
        public void ValidateSettings_FailsOnMissingSelection()
        {
            var settings = new PdfExportSettings
            {
                Scope = PdfExportScope.Selection,
                DestinationPath = @"C:\Exports\Selection.pdf"
            };

            var result = PdfExportEngine.ValidateSettings(settings, new[] { "Sheet1" }, hasSelection: false);
            Assert.IsFalse(result.Success);
            StringAssert.Contains(result.Message, "No active cell range is selected");
        }

        [TestMethod]
        public void ValidateSettings_FailsOnEmptySelectedSheets()
        {
            var settings = new PdfExportSettings
            {
                Scope = PdfExportScope.SelectedSheets,
                SelectedSheets = new List<string>(),
                DestinationPath = @"C:\Exports\Selected.pdf"
            };

            var result = PdfExportEngine.ValidateSettings(settings, new[] { "Sheet1", "Sheet2" });
            Assert.IsFalse(result.Success);
            StringAssert.Contains(result.Message, "select at least one worksheet");
        }

        [TestMethod]
        public void ValidateSettings_FailsOnMissingDestinationPathInSingleMode()
        {
            var settings = new PdfExportSettings
            {
                Scope = PdfExportScope.ActiveSheet,
                OutputMode = PdfOutputMode.SingleCombinedFile,
                DestinationPath = "  "
            };

            var result = PdfExportEngine.ValidateSettings(settings, new[] { "Sheet1" });
            Assert.IsFalse(result.Success);
            StringAssert.Contains(result.Message, "specify a destination PDF file path");
        }

        [TestMethod]
        public void ValidateSettings_FailsOnMissingOutputDirectoryInSeparateMode()
        {
            var settings = new PdfExportSettings
            {
                Scope = PdfExportScope.EntireWorkbook,
                OutputMode = PdfOutputMode.SeparateFilesPerSheet,
                OutputDirectory = ""
            };

            var result = PdfExportEngine.ValidateSettings(settings, new[] { "Sheet1" });
            Assert.IsFalse(result.Success);
            StringAssert.Contains(result.Message, "specify an output folder");
        }

        [TestMethod]
        public void ValidateSettings_FailsOnInvalidPageRange()
        {
            var settings = new PdfExportSettings
            {
                Scope = PdfExportScope.ActiveSheet,
                DestinationPath = @"C:\Exports\Test.pdf",
                ExportPageRange = true,
                PageFrom = 5,
                PageTo = 2
            };

            var result = PdfExportEngine.ValidateSettings(settings, new[] { "Sheet1" });
            Assert.IsFalse(result.Success);
            StringAssert.Contains(result.Message, "End page number cannot be less than start page number");

            settings.PageFrom = 0;
            settings.PageTo = 1;
            var resultZero = PdfExportEngine.ValidateSettings(settings, new[] { "Sheet1" });
            Assert.IsFalse(resultZero.Success);
            StringAssert.Contains(resultZero.Message, "Start page number must be 1 or greater");
        }

        [TestMethod]
        public void MockExcelContext_ExportToPdfWithOptions_SingleFileSuccess()
        {
            var ctx = new MockExcelContext();
            var settings = new PdfExportSettings
            {
                Scope = PdfExportScope.ActiveSheet,
                OutputMode = PdfOutputMode.SingleCombinedFile,
                DestinationPath = @"C:\Exports\ActiveSheet.pdf",
                Orientation = PdfPageOrientation.Landscape,
                PageFit = PdfPageFit.FitAllColumnsOnOnePage,
                PaperSize = PdfPaperSize.A4,
                Quality = PdfQuality.Standard,
                IgnorePrintAreas = true,
                IncludeDocumentProperties = true,
                OpenAfterExport = true,
                OpenOutputFolder = true
            };

            var result = ctx.ExportToPdfWithOptions(settings);

            Assert.IsTrue(result.Success);
            Assert.AreEqual(1, result.FilesExported);
            Assert.AreEqual(@"C:\Exports\ActiveSheet.pdf", result.ExportedFiles[0]);
            Assert.AreEqual(1, ctx.ExportedPdfSettings.Count);
            Assert.AreEqual(PdfPageOrientation.Landscape, ctx.ExportedPdfSettings[0].Orientation);
            Assert.AreEqual(PdfPageFit.FitAllColumnsOnOnePage, ctx.ExportedPdfSettings[0].PageFit);
            Assert.AreEqual(PdfPaperSize.A4, ctx.ExportedPdfSettings[0].PaperSize);
            Assert.IsTrue(ctx.ExportedPdfSettings[0].IgnorePrintAreas);
        }

        [TestMethod]
        public void MockExcelContext_ExportToPdfWithOptions_SeparateFilesMode()
        {
            var ctx = new MockExcelContext();
            var settings = new PdfExportSettings
            {
                Scope = PdfExportScope.SelectedSheets,
                SelectedSheets = new List<string> { "Sheet1", "Sheet2" },
                OutputMode = PdfOutputMode.SeparateFilesPerSheet,
                OutputDirectory = @"C:\Exports\SplitPdfs",
                FileNamePattern = "[Index]_[SheetName]"
            };

            var result = ctx.ExportToPdfWithOptions(settings);

            Assert.IsTrue(result.Success);
            Assert.AreEqual(2, result.FilesExported);
            Assert.AreEqual(Path.Combine(@"C:\Exports\SplitPdfs", "1_Sheet1.pdf"), result.ExportedFiles[0]);
            Assert.AreEqual(Path.Combine(@"C:\Exports\SplitPdfs", "2_Sheet2.pdf"), result.ExportedFiles[1]);
        }

        [TestMethod]
        public void MockExcelContext_ExportToPdfWithOptions_SelectionScope()
        {
            var ctx = new MockExcelContext
            {
                SelectionAddress = "$B$2:$G$25"
            };

            var settings = new PdfExportSettings
            {
                Scope = PdfExportScope.Selection,
                OutputMode = PdfOutputMode.SingleCombinedFile,
                DestinationPath = @"C:\Exports\TableSelection.pdf",
                PageFit = PdfPageFit.FitSheetOnOnePage
            };

            var result = ctx.ExportToPdfWithOptions(settings);

            Assert.IsTrue(result.Success);
            Assert.AreEqual(1, result.FilesExported);
            Assert.AreEqual(@"C:\Exports\TableSelection.pdf", result.ExportedFiles[0]);
        }

        [TestMethod]
        public void LegacyExportToPdf_MaintainsBackwardCompatibility()
        {
            var ctx = new MockExcelContext();
            bool exported = ctx.ExportToPdf(@"C:\Exports\Legacy.pdf", new[] { "Sheet1", "Sheet2" }, openAfterExport: false);

            Assert.IsTrue(exported);
            Assert.AreEqual(1, ctx.ExportedPdfs.Count);
            Assert.AreEqual(@"C:\Exports\Legacy.pdf", ctx.ExportedPdfs[0].dest);
            Assert.IsFalse(ctx.ExportedPdfs[0].open);
            Assert.AreEqual(1, ctx.ExportedPdfSettings.Count);
            Assert.AreEqual(PdfExportScope.SelectedSheets, ctx.ExportedPdfSettings[0].Scope);
        }
    }
}
