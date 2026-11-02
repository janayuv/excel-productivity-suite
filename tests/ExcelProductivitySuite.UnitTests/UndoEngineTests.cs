using System;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Recovery;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class UndoEngineTests
    {
        [TestMethod]
        public void UndoManager_ShouldRecordAndPopUndo()
        {
            var manager = new UndoManager(maxHistory: 10);

            var before1 = new object[,] { { "A1_Old" } };
            var after1 = new object[,] { { "A1_New" } };

            var before2 = new object[,] { { "B1_Old" } };
            var after2 = new object[,] { { "B1_New" } };

            manager.RecordAction("Operation 1", "Sheet1", "$A$1", 1, 1, before1, after1);
            manager.RecordAction("Operation 2", "Sheet1", "$B$1", 1, 2, before2, after2);

            Assert.IsTrue(manager.CanUndo);
            Assert.AreEqual(2, manager.UndoCount);
            Assert.AreEqual("Operation 2", manager.LastUndoDescription);
            Assert.IsFalse(manager.CanRedo);

            var step = manager.PopUndo();
            Assert.IsNotNull(step);
            Assert.AreEqual("Operation 2", step.Description);
            Assert.AreEqual("Sheet1", step.SheetName);
            Assert.AreEqual("$B$1", step.RangeAddress);
            Assert.AreEqual(1, manager.UndoCount);

            manager.PushRedo(step);
            Assert.IsTrue(manager.CanRedo);
            Assert.AreEqual(1, manager.RedoCount);
            Assert.AreEqual("Operation 2", manager.LastRedoDescription);

            var redoStep = manager.PopRedo();
            Assert.IsNotNull(redoStep);
            Assert.AreEqual("Operation 2", redoStep.Description);
            Assert.IsFalse(manager.CanRedo);
        }

        [TestMethod]
        public void UndoManager_ShouldRespectMaxHistory()
        {
            var manager = new UndoManager(maxHistory: 5);

            for (int i = 1; i <= 10; i++)
            {
                var before = new object[,] { { $"Val_{i - 1}" } };
                var after = new object[,] { { $"Val_{i}" } };
                manager.RecordAction($"Op_{i}", "Sheet1", "$A$1", 1, 1, before, after);
            }

            Assert.AreEqual(5, manager.UndoCount);
            Assert.AreEqual("Op_10", manager.LastUndoDescription);

            // Pop all 5
            for (int i = 10; i >= 6; i--)
            {
                var step = manager.PopUndo();
                Assert.IsNotNull(step);
                Assert.AreEqual($"Op_{i}", step.Description);
            }

            Assert.AreEqual(0, manager.UndoCount);
            Assert.IsFalse(manager.CanUndo);
        }

        [TestMethod]
        public void UndoManager_ShouldDiscardNoOpActions()
        {
            var manager = new UndoManager();

            var before = new object[,] { { "Same", 123.45 } };
            var after = new object[,] { { "Same", 123.45 } };

            // Record identical before and after
            manager.RecordAction("No-op Action", "Sheet1", "$A$1:$B$1", 1, 1, before, after);

            Assert.IsFalse(manager.CanUndo);
            Assert.AreEqual(0, manager.UndoCount);
            Assert.IsNull(manager.LastUndoDescription);
        }

        [TestMethod]
        public void UndoManager_ShouldSupportRepeatedUndoAndRedo()
        {
            var manager = new UndoManager();

            var v1 = new object[,] { { "A" } };
            var v2 = new object[,] { { "B" } };
            var v3 = new object[,] { { "C" } };
            var v4 = new object[,] { { "D" } };

            manager.RecordAction("Op 1", "Sheet1", "$A$1", 1, 1, v1, v2);
            manager.RecordAction("Op 2", "Sheet1", "$A$1", 1, 1, v2, v3);
            manager.RecordAction("Op 3", "Sheet1", "$A$1", 1, 1, v3, v4);

            Assert.AreEqual(3, manager.UndoCount);
            Assert.AreEqual(0, manager.RedoCount);

            // Undo Op 3
            var u3 = manager.PopUndo();
            Assert.IsNotNull(u3);
            Assert.AreEqual("Op 3", u3.Description);
            manager.PushRedo(u3);

            // Undo Op 2
            var u2 = manager.PopUndo();
            Assert.IsNotNull(u2);
            Assert.AreEqual("Op 2", u2.Description);
            manager.PushRedo(u2);

            Assert.AreEqual(1, manager.UndoCount);
            Assert.AreEqual(2, manager.RedoCount);
            Assert.AreEqual("Op 2", manager.LastRedoDescription);

            // Redo Op 2
            var r2 = manager.PopRedo();
            Assert.IsNotNull(r2);
            Assert.AreEqual("Op 2", r2.Description);
            manager.PushUndo(r2);

            Assert.AreEqual(2, manager.UndoCount);
            Assert.AreEqual(1, manager.RedoCount);
            Assert.AreEqual("Op 3", manager.LastRedoDescription);

            // Redo Op 3
            var r3 = manager.PopRedo();
            Assert.IsNotNull(r3);
            Assert.AreEqual("Op 3", r3.Description);
            manager.PushUndo(r3);

            Assert.AreEqual(3, manager.UndoCount);
            Assert.AreEqual(0, manager.RedoCount);
            Assert.AreEqual("Op 3", manager.LastUndoDescription);
        }

        [TestMethod]
        public void UndoManager_ShouldDeepCloneSnapshots()
        {
            var mutableBefore = new object[,] { { "Original" } };
            var mutableAfter = new object[,] { { "Modified" } };

            var step = new UndoStep(
                "Cloning Test",
                "Book1.xlsx",
                null,
                "Sheet1",
                "$A$1",
                1,
                1,
                mutableBefore,
                mutableAfter);

            // Mutate the arrays passed to constructor
            mutableBefore[0, 0] = "Tampered Before";
            mutableAfter[0, 0] = "Tampered After";

            // Verify UndoStep retained deep-cloned copies
            Assert.AreEqual("Original", step.BeforeValues[0, 0]);
            Assert.AreEqual("Modified", step.AfterValues[0, 0]);
        }

        [TestMethod]
        public void UndoManager_ShouldClearForSpecificWorkbook()
        {
            var manager = new UndoManager();

            manager.RecordAction("Book1 Op", "Book1.xlsx", @"C:\Book1.xlsx", "Sheet1", "$A$1", 1, 1,
                new object[,] { { 1 } }, new object[,] { { 2 } });
            manager.RecordAction("Book2 Op", "Book2.xlsx", @"C:\Book2.xlsx", "Sheet1", "$A$1", 1, 1,
                new object[,] { { "X" } }, new object[,] { { "Y" } });

            Assert.AreEqual(2, manager.UndoCount);

            // Pop Book2 Op to redo stack
            var popped = manager.PopUndo();
            Assert.IsNotNull(popped);
            manager.PushRedo(popped);

            Assert.AreEqual(1, manager.UndoCount);
            Assert.AreEqual(1, manager.RedoCount);

            // Clear only Book1
            manager.ClearForWorkbook("Book1.xlsx");

            // Undo stack (which had Book1) is now empty
            Assert.AreEqual(0, manager.UndoCount);
            // Redo stack (which had Book2) remains intact
            Assert.AreEqual(1, manager.RedoCount);
            Assert.AreEqual("Book2 Op", manager.LastRedoDescription);

            // Clear Book2
            manager.ClearForWorkbook("Book2.xlsx");
            Assert.AreEqual(0, manager.RedoCount);
            Assert.IsFalse(manager.CanRedo);
        }

        [TestMethod]
        public void UndoManager_100KCellPerformance()
        {
            var manager = new UndoManager(maxHistory: 5);

            const int rows = 1000;
            const int cols = 100;
            var largeBefore = new object[rows, cols];
            var largeAfter = new object[rows, cols];
            var largeFormulas = new object[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    largeBefore[r, c] = r * cols + c;
                    largeAfter[r, c] = (r * cols + c) * 2;
                    if (c == 0)
                    {
                        largeFormulas[r, c] = $"=SUM(B{r + 1}:C{r + 1})";
                    }
                }
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();

            manager.RecordAction(
                "100K Cells Transform",
                "FinancialModel.xlsx",
                @"C:\Models\FinancialModel.xlsx",
                "Sheet1",
                "$A$1:$CV$1000",
                1,
                1,
                largeBefore,
                largeAfter,
                beforeFormulas: largeFormulas);

            sw.Stop();
            long recordMs = sw.ElapsedMilliseconds;

            Assert.IsTrue(manager.CanUndo);
            Assert.AreEqual(1, manager.UndoCount);

            sw.Restart();
            var popped = manager.PopUndo();
            sw.Stop();
            long popMs = sw.ElapsedMilliseconds;

            Assert.IsNotNull(popped);
            Assert.AreEqual(rows, popped.RowCount);
            Assert.AreEqual(cols, popped.ColumnCount);
            Assert.AreEqual("=SUM(B1:C1)", popped.BeforeFormulas?[0, 0]);

            // Assert performance bounds: 100K-cell deep copy and storage should comfortably complete in < 500ms
            Assert.IsTrue(recordMs < 500, $"Snapshot recording for 100,000 cells took {recordMs} ms, expected < 500ms");
            Assert.IsTrue(popMs < 50, $"Popping snapshot for 100,000 cells took {popMs} ms, expected < 50ms");
        }

        [TestMethod]
        public void UndoStep_ShouldPreserveMixedFormulasAndConstants()
        {
            var beforeValues = new object[,]
            {
                { "Label", 100, 200 },
                { "Total", 300, 600 }
            };
            var afterValues = new object[,]
            {
                { "LABEL", 100, 200 },
                { "TOTAL", 300, 600 }
            };
            var beforeFormulas = new object[,]
            {
                { "Label", 100, 200 },
                { "Total", "=B2+C2", "=SUM(B2:C2)*2" }
            };

            var step = new UndoStep(
                "Change Text Case",
                "Report.xlsx",
                null,
                "Summary",
                "$A$1:$C$2",
                1,
                1,
                beforeValues,
                afterValues,
                beforeFormulas: beforeFormulas);

            Assert.AreEqual("Report.xlsx", step.WorkbookName);
            Assert.AreEqual("Summary", step.SheetName);
            Assert.AreEqual(2, step.RowCount);
            Assert.AreEqual(3, step.ColumnCount);

            // Cell [1, 1] is a formula
            Assert.AreEqual("=B2+C2", step.BeforeFormulas?[1, 1]);
            // Cell [0, 1] is a constant number
            Assert.AreEqual(100, step.BeforeFormulas?[0, 1]);
            Assert.AreEqual(100, step.BeforeValues[0, 1]);
        }

        [TestMethod]
        public void UndoManager_NewActionShouldClearRedoStack()
        {
            var manager = new UndoManager();

            var v1 = new object[,] { { 1 } };
            var v2 = new object[,] { { 2 } };
            var v3 = new object[,] { { 3 } };

            manager.RecordAction("Op 1", "Sheet1", "$A$1", 1, 1, v1, v2);

            var popped = manager.PopUndo();
            Assert.IsNotNull(popped);
            manager.PushRedo(popped);
            Assert.IsTrue(manager.CanRedo);

            // New action clears redo
            manager.RecordAction("Op 2", "Sheet1", "$A$1", 1, 1, v2, v3);
            Assert.IsFalse(manager.CanRedo);
            Assert.AreEqual(0, manager.RedoCount);
        }

        [TestMethod]
        public void MockExcelContext_ShouldUndoAndRedoValues()
        {
            var context = new MockExcelContext();

            var v1 = new object[,] { { "Original" } };
            var v2 = new object[,] { { "Modified 1" } };
            var v3 = new object[,] { { "Modified 2" } };

            context.CurrentValues = v1;

            // Record step 1
            context.UndoManager.RecordAction("Change 1", "Sheet1", "$A$1", 1, 1, v1, v2);
            context.CurrentValues = v2;

            // Record step 2
            context.UndoManager.RecordAction("Change 2", "Sheet1", "$A$1", 1, 1, v2, v3);
            context.CurrentValues = v3;

            // Undo once -> should restore Modified 1
            bool undid = context.UndoLastAction();
            Assert.IsTrue(undid);
            Assert.AreEqual("Modified 1", context.CurrentValues[0, 0]);

            // Redo -> should restore Modified 2
            bool redid = context.RedoLastAction();
            Assert.IsTrue(redid);
            Assert.AreEqual("Modified 2", context.CurrentValues[0, 0]);
        }

        [TestMethod]
        public void UndoStep_ShouldPreserveFormulasAndValues()
        {
            var beforeValues = new object[,] { { 10, 20 } };
            var afterValues = new object[,] { { 30, 40 } };
            var beforeFormulas = new object[,] { { "=A1*2", "=B1*2" } };

            var step = new UndoStep(
                "Formula Exact Copy",
                "Sales2026",
                "$C$1:$D$1",
                1,
                3,
                beforeValues,
                afterValues,
                beforeFormulas: beforeFormulas);

            Assert.AreEqual("Formula Exact Copy", step.Description);
            Assert.AreEqual("Sales2026", step.SheetName);
            Assert.AreEqual("$C$1:$D$1", step.RangeAddress);
            Assert.AreEqual(1, step.RowCount);
            Assert.AreEqual(2, step.ColumnCount);
            Assert.AreEqual("=A1*2", step.BeforeFormulas?[0, 0]);
            Assert.AreEqual(30, step.AfterValues[0, 0]);
        }
    }
}
