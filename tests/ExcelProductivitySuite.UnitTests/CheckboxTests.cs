using System;
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
    public class CheckboxTests
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
        public void CheckboxEngine_TryParseRangeAddress_HandlesSingleCell()
        {
            bool success = CheckboxEngine.TryParseRangeAddress("B5", out string? sheet, out int sr, out int sc, out int er, out int ec);
            Assert.IsTrue(success);
            Assert.IsNull(sheet);
            Assert.AreEqual(5, sr);
            Assert.AreEqual(2, sc);
            Assert.AreEqual(5, er);
            Assert.AreEqual(2, ec);
        }

        [TestMethod]
        public void CheckboxEngine_TryParseRangeAddress_HandlesMultiCellRange()
        {
            bool success = CheckboxEngine.TryParseRangeAddress("$A$2:$C$10", out string? sheet, out int sr, out int sc, out int er, out int ec);
            Assert.IsTrue(success);
            Assert.IsNull(sheet);
            Assert.AreEqual(2, sr);
            Assert.AreEqual(1, sc);
            Assert.AreEqual(10, er);
            Assert.AreEqual(3, ec);
        }

        [TestMethod]
        public void CheckboxEngine_TryParseRangeAddress_HandlesSheetQualifiedAddress()
        {
            bool success = CheckboxEngine.TryParseRangeAddress("'Monthly Tasks'!$D$4:$E$8", out string? sheet, out int sr, out int sc, out int er, out int ec);
            Assert.IsTrue(success);
            Assert.AreEqual("Monthly Tasks", sheet);
            Assert.AreEqual(4, sr);
            Assert.AreEqual(4, sc);
            Assert.AreEqual(8, er);
            Assert.AreEqual(5, ec);
        }

        [TestMethod]
        public void CheckboxEngine_TryParseRangeAddress_ReversesInvertedCoordinates()
        {
            bool success = CheckboxEngine.TryParseRangeAddress("D10:B2", out _, out int sr, out int sc, out int er, out int ec);
            Assert.IsTrue(success);
            Assert.AreEqual(2, sr);
            Assert.AreEqual(2, sc);
            Assert.AreEqual(10, er);
            Assert.AreEqual(4, ec);
        }

        [TestMethod]
        public void CheckboxEngine_CalculateCheckboxBounds_HasCaptionExpandsWidth()
        {
            CheckboxEngine.CalculateCheckboxBounds(10.0, 20.0, 100.0, 25.0, hasCaption: true, out double left, out double top, out double width, out double height);
            Assert.AreEqual(12.0, left);
            Assert.IsTrue(top >= 20.0);
            Assert.AreEqual(96.0, width);
            Assert.AreEqual(17.0, height);
        }

        [TestMethod]
        public void CheckboxEngine_CalculateCheckboxBounds_NoCaptionUsesSquareWidth()
        {
            CheckboxEngine.CalculateCheckboxBounds(10.0, 20.0, 100.0, 25.0, hasCaption: false, out double left, out double top, out double width, out double height);
            Assert.AreEqual(12.0, left);
            Assert.AreEqual(16.0, width);
            Assert.AreEqual(17.0, height);
        }

        [TestMethod]
        public void CheckboxEngine_PlanBatchInsert_LinkModeNone_ProducesUnlinkedPlans()
        {
            var data = new object[,]
            {
                { "Task A" },
                { "Task B" },
                { "Task C" }
            };

            var options = new BatchInsertCheckboxOptions
            {
                SourceRangeAddress = "$A$1:$A$3",
                LinkMode = CheckboxLinkMode.None,
                UseCellContentAsLabel = false,
                DisplayLogicalValue = false
            };

            var plans = CheckboxEngine.PlanBatchInsert(data, 1, 1, options);
            Assert.AreEqual(3, plans.Count);
            Assert.IsNull(plans[0].LinkedCellAddress);
            Assert.AreEqual(string.Empty, plans[0].Caption);
            Assert.AreEqual("Task A", plans[0].NewCellValue);
        }

        [TestMethod]
        public void CheckboxEngine_PlanBatchInsert_LinkModeOwnCell_SetsOwnAddressAndLogicalValues()
        {
            var data = new object[,]
            {
                { "Task 1" },
                { "Task 2" }
            };

            var options = new BatchInsertCheckboxOptions
            {
                SourceRangeAddress = "$B$5:$B$6",
                LinkMode = CheckboxLinkMode.OwnCell,
                UseCellContentAsLabel = true,
                DisplayLogicalValue = true,
                InitialChecked = false
            };

            var plans = CheckboxEngine.PlanBatchInsert(data, 5, 2, options);
            Assert.AreEqual(2, plans.Count);
            Assert.AreEqual("$B$5", plans[0].CellAddress);
            Assert.AreEqual("$B$5", plans[0].LinkedCellAddress);
            Assert.AreEqual("Task 1", plans[0].Caption);
            Assert.AreEqual(false, plans[0].NewCellValue);
        }

        [TestMethod]
        public void CheckboxEngine_PlanBatchInsert_LinkModeOtherCells_OffsetsToTargetRange()
        {
            var data = new object[,]
            {
                { "Item 1" },
                { "Item 2" },
                { "Item 3" }
            };

            var options = new BatchInsertCheckboxOptions
            {
                SourceRangeAddress = "$A$2:$A$4",
                LinkMode = CheckboxLinkMode.OtherCells,
                TargetLinkRangeAddress = "$C$2",
                UseCellContentAsLabel = false,
                DisplayLogicalValue = true,
                InitialChecked = true
            };

            var plans = CheckboxEngine.PlanBatchInsert(data, 2, 1, options);
            Assert.AreEqual(3, plans.Count);
            Assert.AreEqual("$A$2", plans[0].CellAddress);
            Assert.AreEqual("$C$2", plans[0].LinkedCellAddress);
            Assert.AreEqual("$A$3", plans[1].CellAddress);
            Assert.AreEqual("$C$3", plans[1].LinkedCellAddress);
            Assert.AreEqual("$A$4", plans[2].CellAddress);
            Assert.AreEqual("$C$4", plans[2].LinkedCellAddress);
            Assert.AreEqual(true, plans[0].NewLinkedCellValue);
        }

        [TestMethod]
        public void CheckboxEngine_PlanBatchInsert_UseCellContentAsLabel_ExtractsCaptionsAndClearsCell()
        {
            var data = new object[,]
            {
                { "Meeting with Team" },
                { "Review budget" }
            };

            var options = new BatchInsertCheckboxOptions
            {
                SourceRangeAddress = "$A$1:$A$2",
                LinkMode = CheckboxLinkMode.None,
                UseCellContentAsLabel = true,
                ClearOriginalContentWhenLabeled = true
            };

            var plans = CheckboxEngine.PlanBatchInsert(data, 1, 1, options);
            Assert.AreEqual(2, plans.Count);
            Assert.AreEqual("Meeting with Team", plans[0].Caption);
            Assert.IsTrue(plans[0].ClearOriginalCell);
            Assert.IsNull(plans[0].NewCellValue);
        }

        [TestMethod]
        public void MockExcelContext_InsertCheckboxes_InsertsExpectedCheckboxes()
        {
            var mock = new MockExcelContext();
            mock.CurrentValues = new object[,]
            {
                { "Alpha" },
                { "Beta" },
                { "Gamma" }
            };
            mock.SelectionStartRow = 1;
            mock.SelectionStartColumn = 1;
            mock.SelectionRowCount = 3;
            mock.SelectionColumnCount = 1;

            var options = new BatchInsertCheckboxOptions
            {
                SourceRangeAddress = "$A$1:$A$3",
                LinkMode = CheckboxLinkMode.OwnCell,
                UseCellContentAsLabel = true,
                DisplayLogicalValue = true,
                InitialChecked = false
            };

            var result = mock.InsertCheckboxes(options);
            Assert.IsTrue(result.Success);
            Assert.AreEqual(3, result.CheckboxesInserted);
            Assert.AreEqual(3, mock.InsertedCheckboxes.Count);
            Assert.AreEqual("$A$1", mock.InsertedCheckboxes[0].CellAddress);
            Assert.AreEqual("Alpha", mock.InsertedCheckboxes[0].Caption);
            Assert.AreEqual(false, mock.CurrentValues[0, 0]);
        }

        [TestMethod]
        public void MockExcelContext_DeleteCheckboxes_SelectionScope_RemovesOnlyIntersecting()
        {
            var mock = new MockExcelContext();
            mock.InsertedCheckboxes.Add(new MockExcelContext.MockCheckboxInfo { Row = 1, Column = 1, CellAddress = "$A$1" });
            mock.InsertedCheckboxes.Add(new MockExcelContext.MockCheckboxInfo { Row = 2, Column = 1, CellAddress = "$A$2" });
            mock.InsertedCheckboxes.Add(new MockExcelContext.MockCheckboxInfo { Row = 5, Column = 1, CellAddress = "$A$5" });

            var delOptions = new BatchDeleteCheckboxOptions
            {
                Scope = CheckboxDeleteScope.Selection,
                RangeAddress = "$A$1:$A$2"
            };

            var res = mock.DeleteCheckboxes(delOptions);
            Assert.IsTrue(res.Success);
            Assert.AreEqual(2, res.CheckboxesDeleted);
            Assert.AreEqual(1, mock.InsertedCheckboxes.Count);
            Assert.AreEqual("$A$5", mock.InsertedCheckboxes[0].CellAddress);
        }

        [TestMethod]
        public void MockExcelContext_DeleteCheckboxes_WorksheetScope_RemovesAll()
        {
            var mock = new MockExcelContext();
            mock.InsertedCheckboxes.Add(new MockExcelContext.MockCheckboxInfo { Row = 1, Column = 1, CellAddress = "$A$1" });
            mock.InsertedCheckboxes.Add(new MockExcelContext.MockCheckboxInfo { Row = 2, Column = 1, CellAddress = "$A$2" });

            var delOptions = new BatchDeleteCheckboxOptions
            {
                Scope = CheckboxDeleteScope.ActiveWorksheet
            };

            var res = mock.DeleteCheckboxes(delOptions);
            Assert.IsTrue(res.Success);
            Assert.AreEqual(2, res.CheckboxesDeleted);
            Assert.AreEqual(0, mock.InsertedCheckboxes.Count);
        }

        [TestMethod]
        public void MockExcelContext_GetCheckboxCount_FiltersByRange()
        {
            var mock = new MockExcelContext();
            mock.InsertedCheckboxes.Add(new MockExcelContext.MockCheckboxInfo { Row = 1, Column = 1, CellAddress = "$A$1" });
            mock.InsertedCheckboxes.Add(new MockExcelContext.MockCheckboxInfo { Row = 2, Column = 1, CellAddress = "$A$2" });
            mock.InsertedCheckboxes.Add(new MockExcelContext.MockCheckboxInfo { Row = 10, Column = 5, CellAddress = "$E$10" });

            Assert.AreEqual(3, mock.GetCheckboxCount(null));
            Assert.AreEqual(2, mock.GetCheckboxCount("$A$1:$A$5"));
            Assert.AreEqual(1, mock.GetCheckboxCount("$E$10"));
            Assert.AreEqual(0, mock.GetCheckboxCount("$B$1:$B$10"));
        }

        [TestMethod]
        public void BatchInsertCheckboxesDialog_InstantiatesAndInitializes()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                mock.SelectionAddress = "$B$2:$B$8";
                var dlg = new BatchInsertCheckboxesDialog(mock);
                Assert.IsNotNull(dlg);
                Assert.AreEqual("$B$2:$B$8", dlg.SourceTextBox.Text);
                Assert.IsTrue(dlg.RadioOwnCellOption.IsChecked == true);
                Assert.IsTrue(dlg.DisplayLogicalValueCheckBox.IsChecked == true);
            });
        }

        [TestMethod]
        public void BatchDeleteCheckboxesDialog_InstantiatesAndInitializes()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                mock.SelectionAddress = "$A$1:$A$10";
                mock.InsertedCheckboxes.Add(new MockExcelContext.MockCheckboxInfo { Row = 1, Column = 1, CellAddress = "$A$1" });
                var dlg = new BatchDeleteCheckboxesDialog(mock);
                Assert.IsNotNull(dlg);
                Assert.AreEqual("$A$1:$A$10", dlg.RangeBox.Text);
                Assert.IsTrue(dlg.RadioSelectionScope.IsChecked == true);
            });
        }

        [TestMethod]
        public void CommandRegistry_FindsCheckboxCommands()
        {
            var registry = new CommandRegistry();

            string addInDllPath = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                @"..\..\..\..\src\ExcelProductivitySuite.AddIn\bin\Debug\ExcelProductivitySuite.AddIn.dll");

            if (!System.IO.File.Exists(addInDllPath))
            {
                addInDllPath = System.IO.Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    @"..\..\..\..\src\ExcelProductivitySuite.AddIn\bin\Release\ExcelProductivitySuite.AddIn.dll");
            }

            if (System.IO.File.Exists(addInDllPath))
            {
                var addInAssembly = System.Reflection.Assembly.LoadFrom(addInDllPath);
                var toolRegistrarType = addInAssembly.GetType("ExcelProductivitySuite.AddIn.Startup.ToolRegistrar");
                if (toolRegistrarType != null)
                {
                    var registerMethod = toolRegistrarType.GetMethod("RegisterAllTools", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    registerMethod?.Invoke(null, new object[] { registry });

                    var insertCmd = registry.GetCommand("CMD_RANGE_INSERT_CHECKBOXES");
                    Assert.IsNotNull(insertCmd, "CMD_RANGE_INSERT_CHECKBOXES should be registered");
                    Assert.AreEqual("Batch Insert Smart Checkboxes", insertCmd.Title);
                    Assert.AreEqual(CommandCategory.RangeAndCells, insertCmd.Category);

                    var deleteCmd = registry.GetCommand("CMD_RANGE_DELETE_CHECKBOXES");
                    Assert.IsNotNull(deleteCmd, "CMD_RANGE_DELETE_CHECKBOXES should be registered");
                    Assert.AreEqual("Batch Delete Checkboxes", deleteCmd.Title);

                    var searchResults = registry.SearchCommands("checkbox");
                    Assert.IsTrue(searchResults.Any(c => c.Id == "CMD_RANGE_INSERT_CHECKBOXES"));
                }
            }
        }
    }
}
