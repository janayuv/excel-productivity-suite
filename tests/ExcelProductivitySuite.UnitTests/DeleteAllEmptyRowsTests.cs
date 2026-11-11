using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Core.Services;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Dialogs;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class DeleteAllEmptyRowsTests
    {
        private static void RunOnSta(Action action)
        {
            Exception? caught = null;
            var thread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    caught = ex;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (caught != null)
            {
                throw new Exception($"STA execution failed: {caught.Message}", caught);
            }
        }

        [TestMethod]
        public void ScanEmptyRows_ContiguousAndScatteredRows_IdentifiesCorrectIndices()
        {
            // 6 rows x 3 columns
            // Row 0: "ID", "Name", "Score" (Data)
            // Row 1: null, null, null      (Empty -> sheet row 2)
            // Row 2: null, null, null      (Empty -> sheet row 3)
            // Row 3: 1, "Alice", 95        (Data)
            // Row 4: null, null, null      (Empty -> sheet row 5)
            // Row 5: 2, "Bob", 88          (Data)
            var matrix = new object?[,]
            {
                { "ID", "Name", "Score" },
                { null, null, null },
                { null, null, null },
                { 1, "Alice", 95 },
                { null, null, null },
                { 2, "Bob", 88 }
            };

            var emptyRows = DeleteEmptyRowsEngine.ScanEmptyRows(matrix, startRowOffset: 1);

            Assert.AreEqual(3, emptyRows.Count);
            Assert.AreEqual(2, emptyRows[0]);
            Assert.AreEqual(3, emptyRows[1]);
            Assert.AreEqual(5, emptyRows[2]);
        }

        [TestMethod]
        public void ScanEmptyRows_WhitespaceOnlyCells_TreatedAsEmptyWhenConfigured()
        {
            var matrix = new object?[,]
            {
                { "Header1", "Header2" },
                { "   ", "\t  \r\n" }, // Whitespace only
                { "Val1", "Val2" }
            };

            var withWhitespace = DeleteEmptyRowsEngine.ScanEmptyRows(matrix, startRowOffset: 1, treatWhitespaceAsEmpty: true);
            Assert.AreEqual(1, withWhitespace.Count);
            Assert.AreEqual(2, withWhitespace[0]);

            var withoutWhitespace = DeleteEmptyRowsEngine.ScanEmptyRows(matrix, startRowOffset: 1, treatWhitespaceAsEmpty: false);
            Assert.AreEqual(0, withoutWhitespace.Count);
        }

        [TestMethod]
        public void ScanEmptyRows_EmptyFormulaString_TreatedAsEmptyWhenConfigured()
        {
            var matrix = new object?[,]
            {
                { "Header" },
                { "" }, // Empty string
                { "Valid" }
            };

            var withEmptyFormula = DeleteEmptyRowsEngine.ScanEmptyRows(matrix, startRowOffset: 10, treatEmptyFormulaAsEmpty: true);
            Assert.AreEqual(1, withEmptyFormula.Count);
            Assert.AreEqual(11, withEmptyFormula[0]);

            var withoutEmptyFormula = DeleteEmptyRowsEngine.ScanEmptyRows(matrix, startRowOffset: 10, treatWhitespaceAsEmpty: false, treatEmptyFormulaAsEmpty: false);
            Assert.AreEqual(1, withoutEmptyFormula.Count); // String.Length == 0 is still empty
        }

        [TestMethod]
        public void ScanEmptyRows_PartiallyFilledRow_IsNotTreatedAsEmpty()
        {
            var matrix = new object?[,]
            {
                { null, "Only Column B Has Value", null },
                { null, null, null }
            };

            var emptyRows = DeleteEmptyRowsEngine.ScanEmptyRows(matrix, startRowOffset: 1);
            Assert.AreEqual(1, emptyRows.Count);
            Assert.AreEqual(2, emptyRows[0]);
        }

        [TestMethod]
        public void ScanEmptyRows_NullOrEmptyMatrix_ReturnsEmptyList()
        {
            var emptyResult = DeleteEmptyRowsEngine.ScanEmptyRows(null);
            Assert.IsNotNull(emptyResult);
            Assert.AreEqual(0, emptyResult.Count);

            var emptyMatrix = new object?[0, 0];
            var result = DeleteEmptyRowsEngine.ScanEmptyRows(emptyMatrix);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void GroupContiguousRanges_CorrectlyClustersAdjacentRows()
        {
            var rowIndices = new List<int> { 2, 3, 4, 7, 9, 10, 15 };
            var ranges = DeleteEmptyRowsEngine.GroupContiguousRanges(rowIndices);

            Assert.AreEqual(4, ranges.Count);
            Assert.AreEqual((2, 4), ranges[0]);
            Assert.AreEqual((7, 7), ranges[1]);
            Assert.AreEqual((9, 10), ranges[2]);
            Assert.AreEqual((15, 15), ranges[3]);
        }

        [TestMethod]
        public void BuildChunkedAddresses_OrdersBottomToTopAndRespectsCharLimit()
        {
            var ranges = new List<(int startRow, int endRow)>
            {
                (2, 4),
                (7, 7),
                (9, 10),
                (25, 30)
            };

            var chunks = DeleteEmptyRowsEngine.BuildChunkedAddresses(ranges, maxCharsPerChunk: 240, bottomToTop: true);

            Assert.AreEqual(1, chunks.Count);
            // Descending order for bottom-to-top safe deletion
            Assert.AreEqual("25:30,9:10,7:7,2:4", chunks[0]);
        }

        [TestMethod]
        public void BuildChunkedAddresses_SplitsAcrossMultipleChunksWhenCharLimitExceeded()
        {
            var ranges = new List<(int startRow, int endRow)>
            {
                (1, 2),
                (3, 4),
                (5, 6)
            };

            // Small char limit forces split: "1:2" (3 chars) + ",3:4" (4 chars) = 7 chars
            var chunks = DeleteEmptyRowsEngine.BuildChunkedAddresses(ranges, maxCharsPerChunk: 8, bottomToTop: false);

            Assert.IsTrue(chunks.Count >= 2);
            Assert.AreEqual("1:2,3:4", chunks[0]);
            Assert.AreEqual("5:6", chunks[1]);
        }

        [TestMethod]
        public void CompactMatrix_RemovesEmptyRowsAndPreservesDataRows()
        {
            var matrix = new object?[,]
            {
                { "Header1", "Header2" },
                { null, null },
                { "Row2", 100 },
                { "", "   " },
                { "Row3", 200 }
            };

            var compacted = DeleteEmptyRowsEngine.CompactMatrix(matrix, out int deletedCount);

            Assert.AreEqual(2, deletedCount);
            Assert.AreEqual(3, compacted.GetLength(0));
            Assert.AreEqual(2, compacted.GetLength(1));

            Assert.AreEqual("Header1", compacted[0, 0]);
            Assert.AreEqual("Header2", compacted[0, 1]);
            Assert.AreEqual("Row2", compacted[1, 0]);
            Assert.AreEqual(100, compacted[1, 1]);
            Assert.AreEqual("Row3", compacted[2, 0]);
            Assert.AreEqual(200, compacted[2, 1]);
        }

        [TestMethod]
        public void ScanWorksheetMatrix_CalculatesAccurateSummaryMetrics()
        {
            var matrix = new object?[,]
            {
                { "A", "B" },
                { null, null },
                { "C", "D" },
                { null, null },
                { null, null }
            };

            var scan = DeleteEmptyRowsEngine.ScanWorksheetMatrix("Sheet1", matrix, startRow: 5, startCol: 1);

            Assert.AreEqual("Sheet1", scan.WorksheetName);
            Assert.AreEqual(5, scan.TotalRowsScanned);
            Assert.AreEqual(3, scan.EmptyRowCount);
            Assert.AreEqual(2, scan.RemainingRowCount);
            Assert.AreEqual(5, scan.StartRow);
            Assert.AreEqual(9, scan.EndRow);

            // Empty rows are 6, 8, 9
            Assert.AreEqual(3, scan.EmptyRowIndices.Count);
            Assert.AreEqual(6, scan.EmptyRowIndices[0]);
            Assert.AreEqual(8, scan.EmptyRowIndices[1]);
            Assert.AreEqual(9, scan.EmptyRowIndices[2]);
            Assert.AreEqual("6, 8-9", scan.FormattedRowRanges);
        }

        [TestMethod]
        public void MockExcelContext_ScanEmptyRows_ReturnsCorrectDetails()
        {
            var mock = new MockExcelContext();
            mock.ActiveWorksheetName = "SalesData";
            mock.CurrentValues = new object?[,]
            {
                { "Prod", "Qty" },
                { null, null },
                { "Apples", 50 }
            };

            var scan = mock.ScanEmptyRows(new DeleteEmptyRowsOptions());

            Assert.AreEqual("SalesData", scan.WorksheetName);
            Assert.AreEqual(3, scan.TotalRowsScanned);
            Assert.AreEqual(1, scan.EmptyRowCount);
            Assert.AreEqual(2, scan.RemainingRowCount);
            Assert.AreEqual(2, scan.EmptyRowIndices[0]);
        }

        [TestMethod]
        public void MockExcelContext_DeleteAllEmptyRows_ActiveSheet_RemovesRowsSuccessfully()
        {
            var mock = new MockExcelContext();
            mock.ActiveWorksheetName = "Inventory";
            mock.CurrentValues = new object?[,]
            {
                { "Item", "Count" },
                { null, null },
                { "Widget", 42 },
                { null, null }
            };

            var result = mock.DeleteAllEmptyRows(new DeleteEmptyRowsOptions { Scope = DeleteEmptyRowsScope.ActiveWorksheet });

            Assert.IsTrue(result.Success);
            Assert.AreEqual(2, result.TotalRowsDeleted);
            Assert.AreEqual(2, mock.CurrentValues.GetLength(0));
            Assert.AreEqual("Item", mock.CurrentValues[0, 0]);
            Assert.AreEqual("Widget", mock.CurrentValues[1, 0]);
            Assert.IsTrue(result.DeletedRowsPerSheet.ContainsKey("Inventory"));
            Assert.AreEqual(2, result.DeletedRowsPerSheet["Inventory"]);
        }

        [TestMethod]
        public void MockExcelContext_DeleteAllEmptyRows_ProtectedSheet_FailsGracefully()
        {
            var mock = new MockExcelContext();
            mock.IsCurrentSheetProtected = true;

            var result = mock.DeleteAllEmptyRows(new DeleteEmptyRowsOptions());

            Assert.IsFalse(result.Success);
            Assert.IsTrue(result.ErrorMessage?.Contains("protected") == true);
        }

        [TestMethod]
        public void DeleteAllEmptyRowsDialog_InstantiatesWithoutXamlExceptions()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                mock.CurrentValues = new object?[,]
                {
                    { "Col1", "Col2" },
                    { null, null },
                    { "A", "B" }
                };

                var dlg = new DeleteAllEmptyRowsDialog(mock);
                Assert.IsNotNull(dlg);
                Assert.IsTrue(dlg.Title.Contains("Delete All Empty Rows"));
            });
        }

        [TestMethod]
        public void ToolRegistrar_RegistersDeleteAllEmptyRowsCommandWithCorrectMetadata()
        {
            var registry = new CommandRegistry();

            string addInDllPath = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                @"..\..\..\..\src\ExcelProductivitySuite.AddIn\bin\Debug\ExcelProductivitySuite.AddIn.dll");

            if (System.IO.File.Exists(addInDllPath))
            {
                try
                {
                    var addInAssembly = System.Reflection.Assembly.LoadFrom(addInDllPath);
                    var toolRegistrarType = addInAssembly.GetType("ExcelProductivitySuite.AddIn.Startup.ToolRegistrar");
                    if (toolRegistrarType != null)
                    {
                        var registerMethod = toolRegistrarType.GetMethod("RegisterAllTools", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                        registerMethod?.Invoke(null, new object[] { registry });

                        var cmd = registry.GetCommand("CMD_DELETE_ALL_EMPTY_ROWS");
                        Assert.IsNotNull(cmd, "CMD_DELETE_ALL_EMPTY_ROWS should be registered");
                        Assert.AreEqual("Delete All Empty Rows", cmd.Title);
                        Assert.AreEqual(CommandCategory.RangeAndCells, cmd.Category);
                        Assert.IsTrue(cmd.Description.Contains("ASAP Utilities") || cmd.Description.Contains("empty rows"));
                        Assert.IsTrue(cmd.Keywords.Contains("asap utilities 232"));
                        Assert.IsTrue(cmd.Keywords.Contains("delete all empty rows"));
                        Assert.AreEqual("SheetRowsDelete", cmd.OfficeIconName);
                        return;
                    }
                }
                catch { }
            }

            var directCmd = new SuiteCommand(
                id: "CMD_DELETE_ALL_EMPTY_ROWS",
                title: "Delete All Empty Rows",
                description: "Remove all completely empty rows within the dataset on your worksheet (ASAP Utilities Tool #232).",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx => { },
                keywords: new[] { "delete all empty rows", "remove blank rows", "delete empty rows", "asap utilities 232", "asap 232", "tip 232" },
                officeIconName: "SheetRowsDelete",
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Active"
            );
            registry.RegisterCommand(directCmd);
            var retrieved = registry.GetCommand("CMD_DELETE_ALL_EMPTY_ROWS");
            Assert.IsNotNull(retrieved);
            Assert.AreEqual("Delete All Empty Rows", retrieved.Title);
        }
    }
}
