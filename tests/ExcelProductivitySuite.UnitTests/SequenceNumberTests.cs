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
            Assert.IsTrue(presets.Count >= 10);
            Assert.IsTrue(presets.Exists(p => p.Name.Contains("Invoice")));
            Assert.IsTrue(presets.Exists(p => p.Name.Contains("Order")));
            Assert.IsTrue(presets.Exists(p => p.Name.Contains("Employee")));
            Assert.IsTrue(presets.Exists(p => p.Name.Contains("Leading Zeros")));
            Assert.IsTrue(presets.Exists(p => p.Type == SequenceType.AlphaUpper));
            Assert.IsTrue(presets.Exists(p => p.Type == SequenceType.RomanUpper));
            Assert.IsTrue(presets.Exists(p => p.Type == SequenceType.CustomList));
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

        [TestMethod]
        public void SequenceType_AlphaUpperAndLower_GeneratesCorrectLetters()
        {
            var defUpper = new SequenceNumberDefinition("AlphaU", 1, 1, 1, "COL-", "", null, 1, false, SequenceType.AlphaUpper);
            Assert.AreEqual("COL-A", defUpper.FormatValue(1));
            Assert.AreEqual("COL-Z", defUpper.FormatValue(26));
            Assert.AreEqual("COL-AA", defUpper.FormatValue(27));
            Assert.AreEqual("COL-AZ", defUpper.FormatValue(52));
            Assert.AreEqual("COL-BA", defUpper.FormatValue(53));

            // Conversion helper
            Assert.AreEqual(1, SequenceNumberDefinition.AlphaToNumber("A"));
            Assert.AreEqual(26, SequenceNumberDefinition.AlphaToNumber("Z"));
            Assert.AreEqual(27, SequenceNumberDefinition.AlphaToNumber("AA"));
            Assert.AreEqual(28, SequenceNumberDefinition.AlphaToNumber("ab"));

            var defLower = new SequenceNumberDefinition("AlphaL", 1, 1, 1, "", "", null, 1, false, SequenceType.AlphaLower);
            var samples = defLower.GeneratePreview(3);
            CollectionAssert.AreEqual(new[] { "a", "b", "c" }, samples);
        }

        [TestMethod]
        public void SequenceType_RomanNumerals_GeneratesCorrectSymbols()
        {
            var defRoman = new SequenceNumberDefinition("RomanU", 1, 1, 1, "SEC-", "", null, 1, false, SequenceType.RomanUpper);
            Assert.AreEqual("SEC-I", defRoman.FormatValue(1));
            Assert.AreEqual("SEC-IV", defRoman.FormatValue(4));
            Assert.AreEqual("SEC-IX", defRoman.FormatValue(9));
            Assert.AreEqual("SEC-XIV", defRoman.FormatValue(14));
            Assert.AreEqual("SEC-XL", defRoman.FormatValue(40));
            Assert.AreEqual("SEC-XC", defRoman.FormatValue(90));
            Assert.AreEqual("SEC-C", defRoman.FormatValue(100));
            Assert.AreEqual("SEC-MCMXCIX", defRoman.FormatValue(1999));

            // Roman to number helper
            Assert.AreEqual(4, SequenceNumberDefinition.RomanToNumber("IV"));
            Assert.AreEqual(9, SequenceNumberDefinition.RomanToNumber("ix"));
            Assert.AreEqual(14, SequenceNumberDefinition.RomanToNumber("XIV"));
            Assert.AreEqual(1999, SequenceNumberDefinition.RomanToNumber("MCMXCIX"));

            var defLowerRoman = new SequenceNumberDefinition("RomanL", 1, 1, 1, "", "", null, 1, false, SequenceType.RomanLower);
            var samples = defLowerRoman.GeneratePreview(4);
            CollectionAssert.AreEqual(new[] { "i", "ii", "iii", "iv" }, samples);
        }

        [TestMethod]
        public void SequenceType_Hexadecimal_GeneratesValidHexCodes()
        {
            var defHex = new SequenceNumberDefinition("HexSeq", 10, 1, 4, "0x", "", null, 1, false, SequenceType.Hexadecimal);
            Assert.AreEqual("0x000A", defHex.FormatValue(10));
            Assert.AreEqual("0x000B", defHex.FormatValue(11));
            Assert.AreEqual("0x000F", defHex.FormatValue(15));
            Assert.AreEqual("0x0010", defHex.FormatValue(16));
            Assert.AreEqual("0x00FF", defHex.FormatValue(255));
        }

        [TestMethod]
        public void SequenceType_CustomList_CyclesThroughConfiguredItems()
        {
            var items = new List<string> { "North", "South", "East", "West" };
            var defCustom = new SequenceNumberDefinition("Compass", 1, 1, 1, "Region: ", "", null, 1, false, SequenceType.CustomList, items);

            var preview = defCustom.GeneratePreview(6);
            CollectionAssert.AreEqual(new[] { "Region: North", "Region: South", "Region: East", "Region: West", "Region: North", "Region: South" }, preview);
        }

        [TestMethod]
        public void DynamicPlaceholders_EvaluatesDateTokensAndSheetName()
        {
            var testDate = new DateTime(2026, 11, 2);
            var def = new SequenceNumberDefinition("DateSeq", 1, 1, 3, "ORD-{YYYY}{MM}-", "-{SHEET}");

            string formatted = def.FormatValue(42, testDate, "Sales2026");
            Assert.AreEqual("ORD-202611-042-Sales2026", formatted);

            // Bracket syntax [YYYY]
            string tokenText = "INV-[YYYY]-[MM]-[DD]";
            string replaced = SequenceNumberDefinition.FormatDynamicTokens(tokenText, testDate, "Test");
            Assert.AreEqual("INV-2026-11-02", replaced);

            // Custom date format {DATE:yyyy-MM-dd}
            string customDateText = "CODE-{DATE:yyyy-MM-dd}";
            string replacedCustom = SequenceNumberDefinition.FormatDynamicTokens(customDateText, testDate, "Test");
            Assert.AreEqual("CODE-2026-11-02", replacedCustom);
        }

        [TestMethod]
        public void FormatMatrix_SerpentineVertical_SnakeColumnOrder()
        {
            // 3 rows x 2 cols
            var matrix = new object?[3, 2];
            var def = new SequenceNumberDefinition("SnakeVert", 1, 1, 1);
            var opts = new SequenceFillOptions { FillOrder = SequenceFillOrder.SerpentineVertical };

            var formatted = SequenceNumberEngine.FormatMatrix(
                matrix,
                def,
                opts,
                out int updated,
                out int skipped,
                out long next);

            Assert.AreEqual(6, updated);
            Assert.AreEqual(7, next);

            // Column 0 goes down (rows 0, 1, 2)
            Assert.AreEqual("1", formatted[0, 0]);
            Assert.AreEqual("2", formatted[1, 0]);
            Assert.AreEqual("3", formatted[2, 0]);

            // Column 1 goes up (rows 2, 1, 0)
            Assert.AreEqual("6", formatted[0, 1]);
            Assert.AreEqual("5", formatted[1, 1]);
            Assert.AreEqual("4", formatted[2, 1]);
        }

        [TestMethod]
        public void FormatMatrix_SerpentineHorizontal_SnakeRowOrder()
        {
            // 2 rows x 3 cols
            var matrix = new object?[2, 3];
            var def = new SequenceNumberDefinition("SnakeHoriz", 1, 1, 1);
            var opts = new SequenceFillOptions { FillOrder = SequenceFillOrder.SerpentineHorizontal };

            var formatted = SequenceNumberEngine.FormatMatrix(
                matrix,
                def,
                opts,
                out int updated,
                out int skipped,
                out long next);

            Assert.AreEqual(6, updated);
            Assert.AreEqual(7, next);

            // Row 0 goes left-to-right (cols 0, 1, 2)
            Assert.AreEqual("1", formatted[0, 0]);
            Assert.AreEqual("2", formatted[0, 1]);
            Assert.AreEqual("3", formatted[0, 2]);

            // Row 1 goes right-to-left (cols 2, 1, 0)
            Assert.AreEqual("6", formatted[1, 0]);
            Assert.AreEqual("5", formatted[1, 1]);
            Assert.AreEqual("4", formatted[1, 2]);
        }

        [TestMethod]
        public void FormatMatrix_RestartOnValueChange_GroupedNumbering()
        {
            // 5 rows x 2 cols: Col 0 is Category, Col 1 is Sequence target
            var matrix = new object?[5, 2]
            {
                { "Apples", null },
                { "Apples", null },
                { "Apples", null },
                { "Oranges", null },
                { "Oranges", null }
            };

            var def = new SequenceNumberDefinition("GroupSeq", 1, 1, 1);
            var opts = new SequenceFillOptions
            {
                FillOrder = SequenceFillOrder.Vertical,
                RestartOnValueChange = true,
                RestartKeyColumnOffset = -1 // inspect column to the left (Col 0)
            };

            var formatted = SequenceNumberEngine.FormatMatrix(
                matrix,
                def,
                opts,
                out int updated,
                out int skipped,
                out _);

            Assert.AreEqual(10, updated); // Processes all cells in both cols
            // Let's test on just Col 1:
            var targetCol = new object?[5, 2]
            {
                { "CatA", null },
                { "CatA", null },
                { "CatB", null },
                { "CatB", null },
                { "CatC", null }
            };

            var optsTarget = new SequenceFillOptions
            {
                FillOrder = SequenceFillOrder.Vertical,
                OnlyFillBlankCells = true,
                RestartOnValueChange = true,
                RestartKeyColumnOffset = -1
            };

            var res = SequenceNumberEngine.FormatMatrix(targetCol, def, optsTarget, out int u, out int s, out _);
            Assert.AreEqual(5, u);
            Assert.AreEqual("1", res[0, 1]);
            Assert.AreEqual("2", res[1, 1]);
            Assert.AreEqual("1", res[2, 1]); // Restart on CatB
            Assert.AreEqual("2", res[3, 1]);
            Assert.AreEqual("1", res[4, 1]); // Restart on CatC
        }

        [TestMethod]
        public void FormatMatrix_RandomizeOrder_GeneratesUniqueShuffledNumbers()
        {
            var matrix = new object?[10, 1];
            var def = new SequenceNumberDefinition("RandomSeq", 1, 1, 2);
            var opts = new SequenceFillOptions
            {
                FillOrder = SequenceFillOrder.Vertical,
                RandomizeOrder = true
            };

            var formatted = SequenceNumberEngine.FormatMatrix(
                matrix,
                def,
                opts,
                out int updated,
                out int skipped,
                out long next);

            Assert.AreEqual(10, updated);
            Assert.AreEqual(11, next);

            var collected = new HashSet<string>();
            for (int r = 0; r < 10; r++)
            {
                string val = formatted[r, 0]?.ToString() ?? "";
                Assert.IsFalse(string.IsNullOrEmpty(val));
                collected.Add(val);
            }

            // All 10 numbers from "01" to "10" must be present without duplicates
            Assert.AreEqual(10, collected.Count);
            for (int i = 1; i <= 10; i++)
            {
                Assert.IsTrue(collected.Contains(i.ToString("D2")));
            }
        }

        [TestMethod]
        public void FormatMatrix_StepRows_FillsEveryNthRow()
        {
            var matrix = new object?[6, 1];
            var def = new SequenceNumberDefinition("StepRowSeq", 10, 10, 1);
            var opts = new SequenceFillOptions
            {
                FillOrder = SequenceFillOrder.Vertical,
                StepRows = 2 // Row 0, Row 2, Row 4
            };

            var formatted = SequenceNumberEngine.FormatMatrix(
                matrix,
                def,
                opts,
                out int updated,
                out int skipped,
                out long next);

            Assert.AreEqual(3, updated);
            Assert.AreEqual(3, skipped);
            Assert.AreEqual(40, next);

            Assert.AreEqual("10", formatted[0, 0]);
            Assert.IsNull(formatted[1, 0]);
            Assert.AreEqual("20", formatted[2, 0]);
            Assert.IsNull(formatted[3, 0]);
            Assert.AreEqual("30", formatted[4, 0]);
            Assert.IsNull(formatted[5, 0]);
        }

        [TestMethod]
        public void SequenceNumberEngine_ExportAndImport_PersistsDefinitionsAccurately()
        {
            string tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"eps_seq_test_{Guid.NewGuid():N}.json");

            try
            {
                var original = new List<SequenceNumberDefinition>
                {
                    new SequenceNumberDefinition("TestAlpha", 1, 2, 2, "REV-", "-END", null, 1, false, SequenceType.AlphaUpper),
                    new SequenceNumberDefinition("TestCustom", 1, 1, 1, "", "", null, 1, false, SequenceType.CustomList,
                        new List<string> { "P1", "P2", "P3" }, useDynamicPlaceholders: true, outputAsNumeric: false)
                };

                bool exportSuccess = SequenceNumberEngine.ExportToFile(tempFile, original, out string? exportErr);
                Assert.IsTrue(exportSuccess);
                Assert.IsNull(exportErr);

                bool importSuccess = SequenceNumberEngine.ImportFromFile(tempFile, out var imported, out string? importErr);
                Assert.IsTrue(importSuccess);
                Assert.IsNull(importErr);
                Assert.AreEqual(2, imported.Count);

                Assert.AreEqual("TestAlpha", imported[0].Name);
                Assert.AreEqual(SequenceType.AlphaUpper, imported[0].Type);
                Assert.AreEqual("REV-", imported[0].Prefix);
                Assert.AreEqual("-END", imported[0].Suffix);
                Assert.AreEqual(2, imported[0].Increment);

                Assert.AreEqual("TestCustom", imported[1].Name);
                Assert.AreEqual(SequenceType.CustomList, imported[1].Type);
                Assert.AreEqual(3, imported[1].CustomListItems.Count);
                Assert.AreEqual("P1", imported[1].CustomListItems[0]);
                Assert.AreEqual("P2", imported[1].CustomListItems[1]);
                Assert.AreEqual("P3", imported[1].CustomListItems[2]);
            }
            finally
            {
                if (System.IO.File.Exists(tempFile))
                {
                    System.IO.File.Delete(tempFile);
                }
            }
        }
    }
}
