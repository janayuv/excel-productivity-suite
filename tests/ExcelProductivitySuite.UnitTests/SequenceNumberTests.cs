using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Ranges;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class SequenceNumberTests
    {
        [TestMethod]
        public void FormatValue_CorrectlyAppliesPrefixPaddingAndSuffix()
        {
            var def = new SequenceNumberDefinition("Test", 1, 1, 4, "INV-", "-2026");
            Assert.AreEqual("INV-0001-2026", def.FormatValue(1));
            Assert.AreEqual("INV-0042-2026", def.FormatValue(42));
            Assert.AreEqual("INV-1000-2026", def.FormatValue(1000));

            // Without padding
            var def2 = new SequenceNumberDefinition("Simple", 5, 1, 1, "", "");
            Assert.AreEqual("5", def2.FormatValue(5));
            Assert.AreEqual("99", def2.FormatValue(99));

            // Negative numbers
            var def3 = new SequenceNumberDefinition("Negative", -1, 1, 3, "ID-");
            Assert.AreEqual("ID--005", def3.FormatValue(-5));
        }

        [TestMethod]
        public void GeneratePreview_HandlesStepAndEndNumberLoop()
        {
            // Stepping by 5
            var defStep = new SequenceNumberDefinition("Step5", 10, increment: 5, digits: 2);
            var preview = defStep.GeneratePreview(4);
            CollectionAssert.AreEqual(new[] { "10", "15", "20", "25" }, preview);

            // Cyclic sequence 1 to 3
            var defCycle = new SequenceNumberDefinition("Cycle", 1, increment: 1, digits: 1, prefix: "A", endNumber: 3);
            var cyclePreview = defCycle.GeneratePreview(7);
            CollectionAssert.AreEqual(new[] { "A1", "A2", "A3", "A1", "A2", "A3", "A1" }, cyclePreview);

            // Repeat count: repeat each number 2 times
            var defRepeat = new SequenceNumberDefinition("Pairs", 1, increment: 1, digits: 1, repeatCount: 2);
            var repeatPreview = defRepeat.GeneratePreview(6);
            CollectionAssert.AreEqual(new[] { "1", "1", "2", "2", "3", "3" }, repeatPreview);
        }

        [TestMethod]
        public void GenerateSequence_OutputsCorrectValuesAndNextNumber()
        {
            var def = new SequenceNumberDefinition("Seq", 100, 10, 4, "ORD-");
            var result = SequenceNumberEngine.GenerateSequence(def, 3, out long next);

            Assert.AreEqual(3, result.Count);
            Assert.AreEqual("ORD-0100", result[0]);
            Assert.AreEqual("ORD-0110", result[1]);
            Assert.AreEqual("ORD-0120", result[2]);
            Assert.AreEqual(130, next);

            // Continue from current
            def.CurrentNumber = 150;
            var resultCont = SequenceNumberEngine.GenerateSequence(def, 2, continueFromCurrent: true, out long nextCont);
            Assert.AreEqual("ORD-0150", resultCont[0]);
            Assert.AreEqual("ORD-0160", resultCont[1]);
            Assert.AreEqual(170, nextCont);
        }

        [TestMethod]
        public void FormatMatrix_VerticalFill_FillsColumnsFirst()
        {
            // 3 rows x 2 cols matrix
            var matrix = new object?[3, 2];
            var def = new SequenceNumberDefinition("Vert", 1, 1, 2);
            var opts = new SequenceFillOptions { FillOrder = SequenceFillOrder.Vertical };

            var formatted = SequenceNumberEngine.FormatMatrix(
                matrix,
                def,
                opts,
                out int updated,
                out int skipped,
                out long next);

            Assert.AreEqual(6, updated);
            Assert.AreEqual(0, skipped);
            Assert.AreEqual(7, next);

            // Column 0: rows 0, 1, 2 -> "01", "02", "03"
            Assert.AreEqual("01", formatted[0, 0]);
            Assert.AreEqual("02", formatted[1, 0]);
            Assert.AreEqual("03", formatted[2, 0]);

            // Column 1: rows 0, 1, 2 -> "04", "05", "06"
            Assert.AreEqual("04", formatted[0, 1]);
            Assert.AreEqual("05", formatted[1, 1]);
            Assert.AreEqual("06", formatted[2, 1]);
        }

        [TestMethod]
        public void FormatMatrix_HorizontalFill_FillsRowsFirst()
        {
            var matrix = new object?[2, 3];
            var def = new SequenceNumberDefinition("Horiz", 1, 1, 2);
            var opts = new SequenceFillOptions { FillOrder = SequenceFillOrder.Horizontal };

            var formatted = SequenceNumberEngine.FormatMatrix(
                matrix,
                def,
                opts,
                out int updated,
                out int skipped,
                out long next);

            Assert.AreEqual(6, updated);
            Assert.AreEqual(0, skipped);
            Assert.AreEqual(7, next);

            // Row 0: cols 0, 1, 2 -> "01", "02", "03"
            Assert.AreEqual("01", formatted[0, 0]);
            Assert.AreEqual("02", formatted[0, 1]);
            Assert.AreEqual("03", formatted[0, 2]);

            // Row 1: cols 0, 1, 2 -> "04", "05", "06"
            Assert.AreEqual("04", formatted[1, 0]);
            Assert.AreEqual("05", formatted[1, 1]);
            Assert.AreEqual("06", formatted[1, 2]);
        }

        [TestMethod]
        public void FormatMatrix_OnlyFillBlankCells_SkipsExistingData()
        {
            var matrix = new object?[3, 1]
            {
                { "ExistingHeader" },
                { null },
                { "KeepMe" }
            };

            var def = new SequenceNumberDefinition("BlanksOnly", 1, 1, 1);
            var opts = new SequenceFillOptions
            {
                FillOrder = SequenceFillOrder.Vertical,
                OnlyFillBlankCells = true
            };

            var formatted = SequenceNumberEngine.FormatMatrix(
                matrix,
                def,
                opts,
                out int updated,
                out int skipped,
                out long next);

            Assert.AreEqual(1, updated);
            Assert.AreEqual(2, skipped);
            Assert.AreEqual(2, next);

            Assert.AreEqual("ExistingHeader", formatted[0, 0]);
            Assert.AreEqual("1", formatted[1, 0]);
            Assert.AreEqual("KeepMe", formatted[2, 0]);
        }

        [TestMethod]
        public void DefaultPresets_ArePopulatedWithExpectedTemplates()
        {
            var presets = SequenceNumberEngine.GetDefaultPresets();
            Assert.IsTrue(presets.Count >= 7);
            Assert.IsTrue(presets.Exists(p => p.Name.Contains("Invoice")));
            Assert.IsTrue(presets.Exists(p => p.Name.Contains("Order")));
            Assert.IsTrue(presets.Exists(p => p.Name.Contains("Employee")));
            Assert.IsTrue(presets.Exists(p => p.Name.Contains("Leading Zeros")));
        }

        [TestMethod]
        public void MockExcelContext_InsertSequenceNumbers_ExecutesAndMaintainsState()
        {
            var ctx = new MockExcelContext();
            ctx.CurrentValues = new object?[3, 1];

            var def = new SequenceNumberDefinition("MockTest", 1, 1, 3, "ID-");
            var opts = new SequenceFillOptions { FillOrder = SequenceFillOrder.Vertical, ContinueFromLast = true };

            var result = ctx.InsertSequenceNumbers(def, opts);
            Assert.IsTrue(result.Success);
            Assert.AreEqual(3, result.AffectedCellsCount);
            Assert.AreEqual("ID-001", ctx.CurrentValues[0, 0]);
            Assert.AreEqual("ID-002", ctx.CurrentValues[1, 0]);
            Assert.AreEqual("ID-003", ctx.CurrentValues[2, 0]);
            Assert.AreEqual(4, def.CurrentNumber);

            // Second insertion continues from 4
            ctx.CurrentValues = new object?[2, 1];
            var result2 = ctx.InsertSequenceNumbers(def, opts);
            Assert.IsTrue(result2.Success);
            Assert.AreEqual("ID-004", ctx.CurrentValues[0, 0]);
            Assert.AreEqual("ID-005", ctx.CurrentValues[1, 0]);
            Assert.AreEqual(6, def.CurrentNumber);

            // Reset counter
            def.Reset();
            Assert.AreEqual(1, def.CurrentNumber);
        }
    }
}
