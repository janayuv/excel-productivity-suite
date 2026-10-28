using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Core.Services;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.Engine.Analysis;
using ExcelProductivitySuite.Engine.Files;
using ExcelProductivitySuite.Engine.Formulas;
using ExcelProductivitySuite.Engine.Formatting;
using ExcelProductivitySuite.Engine.Navigation;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.Engine.Recovery;
using ExcelProductivitySuite.Engine.Worksheets;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class RangeAndSheetTests
    {
        [TestMethod]
        public void FillBlanks_Down_FillsFromAbove()
        {
            // 3 rows x 2 cols (0-based)
            object?[,] matrix = new object?[3, 2]
            {
                { "Header1", "Header2" },
                { null,      "Val2" },
                { null,      null }
            };

            var options = new FillBlankOptions { Direction = FillBlankDirection.Down };
            var result = RangeFill.FillBlanks(matrix, options, out int count);

            Assert.AreEqual(3, count);
            Assert.AreEqual("Header1", result[1, 0]);
            Assert.AreEqual("Header1", result[2, 0]);
            Assert.AreEqual("Val2", result[2, 1]);
        }

        [TestMethod]
        public void FillBlanks_FixedValue_FillsAllBlanks()
        {
            object?[,] matrix = new object?[2, 2]
            {
                { "Item", null },
                { "",     "Value" }
            };

            var options = new FillBlankOptions
            {
                Direction = FillBlankDirection.FixedValue,
                FixedValue = "N/A"
            };

            var result = RangeFill.FillBlanks(matrix, options, out int count);
            Assert.AreEqual(2, count);
            Assert.AreEqual("N/A", result[0, 1]);
            Assert.AreEqual("N/A", result[1, 0]);
        }

        [TestMethod]
        public void DuplicateFinder_IdentifiesDuplicateRows()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { "A", "10" },
                { "B", "20" },
                { "A", "10" }, // duplicate of row 0
                { "C", "30" }
            };

            var options = new DuplicateDetectionOptions
            {
                KeepStrategy = DuplicateKeepStrategy.KeepFirst,
                CompareEntireRows = true
            };

            bool[] isDuplicate = DuplicateFinder.IdentifyDuplicateRows(matrix, options);

            Assert.AreEqual(4, isDuplicate.Length);
            Assert.IsFalse(isDuplicate[0]);
            Assert.IsFalse(isDuplicate[1]);
            Assert.IsTrue(isDuplicate[2]);
            Assert.IsFalse(isDuplicate[3]);
        }

        [TestMethod]
        public void WorksheetNameValidator_DetectsInvalidNames()
        {
            var r1 = WorksheetNameValidator.ValidateSheetName("");
            Assert.IsFalse(r1.IsValid);

            var r2 = WorksheetNameValidator.ValidateSheetName("Sheet:Invalid");
            Assert.IsFalse(r2.IsValid);

            var r3 = WorksheetNameValidator.ValidateSheetName(new string('A', 35));
            Assert.IsFalse(r3.IsValid);

            var r4 = WorksheetNameValidator.ValidateSheetName("History");
            Assert.IsFalse(r4.IsValid);

            var r5 = WorksheetNameValidator.ValidateSheetName("ValidSheet_1");
            Assert.IsTrue(r5.IsValid);
        }

        [TestMethod]
        public void BatchRename_GeneratesUniqueSanitizedNames()
        {
            var oldNames = new List<string> { "Data", "Report", "Sales" };
            var options = new BatchRenameOptions
            {
                Prefix = "Q3_",
                Suffix = "_2026",
                AppendSequenceNumber = true,
                SequenceStart = 1,
                SequenceFormat = "D2"
            };

            var renamed = WorksheetNameValidator.GenerateRenamedNames(oldNames, options);

            Assert.AreEqual("Q3_Data_2026_01", renamed["Data"]);
            Assert.AreEqual("Q3_Report_2026_02", renamed["Report"]);
            Assert.AreEqual("Q3_Sales_2026_03", renamed["Sales"]);
        }

        // =========================================================================
        // Batch 3: Range and Cell Operations Unit Tests
        // =========================================================================

        #region RangeTransposer Tests

        [TestMethod]
        public void RangeTransposer_RectangularMatrix_SwapsDimensionsAndValues()
        {
            // 2 rows x 3 cols
            object?[,] matrix = new object?[2, 3]
            {
                { "A1", "B1", "C1" },
                { "A2", "B2", "C2" }
            };

            var transposed = RangeTransposer.Transpose(matrix);

            // New dimensions: 3 rows x 2 cols
            Assert.AreEqual(3, transposed.GetLength(0));
            Assert.AreEqual(2, transposed.GetLength(1));

            Assert.AreEqual("A1", transposed[0, 0]);
            Assert.AreEqual("A2", transposed[0, 1]);
            Assert.AreEqual("B1", transposed[1, 0]);
            Assert.AreEqual("B2", transposed[1, 1]);
            Assert.AreEqual("C1", transposed[2, 0]);
            Assert.AreEqual("C2", transposed[2, 1]);
        }

        #endregion

        #region RangeFlipper Tests

        [TestMethod]
        public void RangeFlipper_Horizontal_ReversesColumnOrder()
        {
            object?[,] matrix = new object?[2, 3]
            {
                { "Col1", "Col2", "Col3" },
                { "10",   "20",   "30" }
            };

            var flipped = RangeFlipper.FlipHorizontal(matrix);

            Assert.AreEqual("Col3", flipped[0, 0]);
            Assert.AreEqual("Col2", flipped[0, 1]);
            Assert.AreEqual("Col1", flipped[0, 2]);

            Assert.AreEqual("30", flipped[1, 0]);
            Assert.AreEqual("20", flipped[1, 1]);
            Assert.AreEqual("10", flipped[1, 2]);
        }

        [TestMethod]
        public void RangeFlipper_Vertical_ReversesRowOrder()
        {
            object?[,] matrix = new object?[3, 2]
            {
                { "Row1", "A" },
                { "Row2", "B" },
                { "Row3", "C" }
            };

            var flipped = RangeFlipper.FlipVertical(matrix);

            Assert.AreEqual("Row3", flipped[0, 0]);
            Assert.AreEqual("C", flipped[0, 1]);
            Assert.AreEqual("Row2", flipped[1, 0]);
            Assert.AreEqual("B", flipped[1, 1]);
            Assert.AreEqual("Row1", flipped[2, 0]);
            Assert.AreEqual("A", flipped[2, 1]);
        }

        #endregion

        #region RangeSwapper Tests

        [TestMethod]
        public void RangeSwapper_ExchangesDataBetweenRanges()
        {
            object?[,] rangeA = new object?[2, 2] { { "A1", "A2" }, { "A3", "A4" } };
            object?[,] rangeB = new object?[2, 2] { { "B1", "B2" }, { "B3", "B4" } };

            var (newA, newB) = RangeSwapper.Swap(rangeA, rangeB, requireEqualDimensions: true);

            Assert.AreEqual("B1", newA[0, 0]);
            Assert.AreEqual("A1", newB[0, 0]);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void RangeSwapper_MismatchedDimensions_ThrowsWhenRequired()
        {
            object?[,] rangeA = new object?[2, 2];
            object?[,] rangeB = new object?[3, 3];

            RangeSwapper.Swap(rangeA, rangeB, requireEqualDimensions: true);
        }

        #endregion

        #region SignChanger Tests

        [TestMethod]
        public void SignChanger_InvertSign_TogglesNumbers()
        {
            var options = new SignChangeOptions { Mode = SignChangeMode.InvertSign };

            Assert.AreEqual(-50.0, (double)SignChanger.ChangeSign(50.0, options)!, 0.0001);
            Assert.AreEqual(25.5, (double)SignChanger.ChangeSign(-25.5, options)!, 0.0001);
            Assert.AreEqual(0.0, (double)SignChanger.ChangeSign(0.0, options)!, 0.0001);
        }

        [TestMethod]
        public void SignChanger_MakeAllPositive_ReturnsAbsoluteValue()
        {
            var options = new SignChangeOptions { Mode = SignChangeMode.MakeAllPositive };

            Assert.AreEqual(100.0, (double)SignChanger.ChangeSign(-100.0, options)!, 0.0001);
            Assert.AreEqual(42.0, (double)SignChanger.ChangeSign(42.0, options)!, 0.0001);
        }

        [TestMethod]
        public void SignChanger_MakeAllNegative_ReturnsNegativeValue()
        {
            var options = new SignChangeOptions { Mode = SignChangeMode.MakeAllNegative };

            Assert.AreEqual(-100.0, (double)SignChanger.ChangeSign(100.0, options)!, 0.0001);
            Assert.AreEqual(-42.0, (double)SignChanger.ChangeSign(-42.0, options)!, 0.0001);
        }

        [TestMethod]
        public void SignChanger_Range_UpdatesNumericCellsAndPreservesStrings()
        {
            object?[,] matrix = new object?[2, 2]
            {
                { "Label", 50.0 },
                { -12.0,  null }
            };

            var options = new SignChangeOptions { Mode = SignChangeMode.InvertSign };
            var result = SignChanger.ChangeSignRange(matrix, options, out int affected);

            Assert.AreEqual(2, affected);
            Assert.AreEqual("Label", result[0, 0]);
            Assert.AreEqual(-50.0, (double)result[0, 1]!, 0.0001);
            Assert.AreEqual(12.0, (double)result[1, 0]!, 0.0001);
        }

        #endregion

        #region ValueRounder Tests

        [TestMethod]
        public void ValueRounder_StandardRounding_RoundsAwayFromZero()
        {
            var options = new ValueRoundOptions { Mode = RoundingMode.Standard, DecimalPlaces = 2 };

            Assert.AreEqual(12.35, (double)ValueRounder.Round(12.3456, options)!, 0.0001);
            Assert.AreEqual(10.0, (double)ValueRounder.Round(9.999, options)!, 0.0001);
        }

        [TestMethod]
        public void ValueRounder_RoundUpCeiling_RoundsUpward()
        {
            var options = new ValueRoundOptions { Mode = RoundingMode.RoundUpCeiling, DecimalPlaces = 2 };

            Assert.AreEqual(12.35, (double)ValueRounder.Round(12.341, options)!, 0.0001);
        }

        [TestMethod]
        public void ValueRounder_RoundDownFloor_RoundsDownward()
        {
            var options = new ValueRoundOptions { Mode = RoundingMode.RoundDownFloor, DecimalPlaces = 2 };

            Assert.AreEqual(12.34, (double)ValueRounder.Round(12.349, options)!, 0.0001);
        }

        [TestMethod]
        public void ValueRounder_Truncate_StripsFractionalDigits()
        {
            var options = new ValueRoundOptions { Mode = RoundingMode.Truncate, DecimalPlaces = 1 };

            Assert.AreEqual(45.6, (double)ValueRounder.Round(45.678, options)!, 0.0001);
        }

        [TestMethod]
        public void ValueRounder_NearestMultiple_SnapsToSpecifiedMultiple()
        {
            var options = new ValueRoundOptions { Mode = RoundingMode.NearestMultiple, Multiple = 5.0 };

            Assert.AreEqual(25.0, (double)ValueRounder.Round(23.0, options)!, 0.0001);
            Assert.AreEqual(20.0, (double)ValueRounder.Round(22.0, options)!, 0.0001);
            Assert.AreEqual(50.0, (double)ValueRounder.Round(48.5, options)!, 0.0001);
        }

        #endregion

        #region RangeRandomizer Tests

        [TestMethod]
        public void RangeRandomizer_EntireRows_PreservesRecordAssociation()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { "Alice", 100 },
                { "Bob",   200 },
                { "Carol", 300 },
                { "Dave",  400 }
            };

            var options = new RandomSortOptions { Mode = RandomSortMode.EntireRows, HasHeaderRow = false };
            var result = RangeRandomizer.Randomize(matrix, options, new Random(12345));

            Assert.AreEqual(4, result.GetLength(0));
            Assert.AreEqual(2, result.GetLength(1));

            // Verify that for each row, the name and score are still together!
            for (int r = 0; r < 4; r++)
            {
                string name = result[r, 0]?.ToString() ?? "";
                int score = Convert.ToInt32(result[r, 1]);

                if (name == "Alice") Assert.AreEqual(100, score);
                else if (name == "Bob") Assert.AreEqual(200, score);
                else if (name == "Carol") Assert.AreEqual(300, score);
                else if (name == "Dave") Assert.AreEqual(400, score);
            }
        }

        [TestMethod]
        public void RangeRandomizer_HasHeaderRow_ExcludesFirstRow()
        {
            object?[,] matrix = new object?[3, 2]
            {
                { "Name", "Score" },
                { "A", 1 },
                { "B", 2 }
            };

            var options = new RandomSortOptions { Mode = RandomSortMode.EntireRows, HasHeaderRow = true };
            var result = RangeRandomizer.Randomize(matrix, options, new Random(42));

            Assert.AreEqual("Name", result[0, 0]);
            Assert.AreEqual("Score", result[0, 1]);
        }

        #endregion

        #region SafeRangeMerger Tests

        [TestMethod]
        public void SafeRangeMerger_RowByRow_CombinesTextAndClearsSiblings()
        {
            object?[,] matrix = new object?[2, 3]
            {
                { "First", "Middle", "Last" },
                { "John",  null,     "Smith" }
            };

            var options = new SafeMergeOptions
            {
                Direction = SafeMergeDirection.RowByRow,
                Separator = " "
            };

            var result = SafeRangeMerger.PrepareMerge(matrix, options, out int groups);

            Assert.AreEqual(2, groups);
            Assert.AreEqual("First Middle Last", result[0, 0]);
            Assert.IsNull(result[0, 1]);
            Assert.IsNull(result[0, 2]);

            Assert.AreEqual("John Smith", result[1, 0]);
            Assert.IsNull(result[1, 1]);
            Assert.IsNull(result[1, 2]);
        }

        [TestMethod]
        public void SafeRangeMerger_EntireSelection_CombinesToTopLeft()
        {
            object?[,] matrix = new object?[2, 2]
            {
                { "A", "B" },
                { "C", "D" }
            };

            var options = new SafeMergeOptions
            {
                Direction = SafeMergeDirection.EntireSelection,
                Separator = ", "
            };

            var result = SafeRangeMerger.PrepareMerge(matrix, options, out int groups);

            Assert.AreEqual(1, groups);
            Assert.AreEqual("A, B, C, D", result[0, 0]);
            Assert.IsNull(result[0, 1]);
            Assert.IsNull(result[1, 0]);
            Assert.IsNull(result[1, 1]);
        }

        #endregion

        #region BlankRowManager Tests

        [TestMethod]
        public void BlankRowManager_InsertBlankRows_EveryNRows()
        {
            object?[,] matrix = new object?[3, 2]
            {
                { "R1C1", "R1C2" },
                { "R2C1", "R2C2" },
                { "R3C1", "R3C2" }
            };

            var options = new BlankRowInsertOptions
            {
                Mode = BlankRowInsertMode.EveryNRows,
                Interval = 1,
                BlankRowCount = 1,
                HasHeaderRow = false
            };

            var result = BlankRowManager.InsertBlankRows(matrix, options);

            // 3 data rows with 1 blank after each of the first 2 = 5 rows total
            Assert.AreEqual(5, result.GetLength(0));
            Assert.AreEqual("R1C1", result[0, 0]);
            Assert.IsNull(result[1, 0]);
            Assert.AreEqual("R2C1", result[2, 0]);
            Assert.IsNull(result[3, 0]);
            Assert.AreEqual("R3C1", result[4, 0]);
        }

        [TestMethod]
        public void BlankRowManager_InsertBlankRows_OnValueChange()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { "Group A", 10 },
                { "Group A", 20 },
                { "Group B", 30 },
                { "Group B", 40 }
            };

            var options = new BlankRowInsertOptions
            {
                Mode = BlankRowInsertMode.OnValueChange,
                KeyColumnIndex = 0,
                BlankRowCount = 1,
                HasHeaderRow = false
            };

            var result = BlankRowManager.InsertBlankRows(matrix, options);

            // Group A (2 rows), 1 blank row, Group B (2 rows) = 5 rows total
            Assert.AreEqual(5, result.GetLength(0));
            Assert.AreEqual("Group A", result[0, 0]);
            Assert.AreEqual("Group A", result[1, 0]);
            Assert.IsNull(result[2, 0]);
            Assert.AreEqual("Group B", result[3, 0]);
            Assert.AreEqual("Group B", result[4, 0]);
        }

        [TestMethod]
        public void BlankRowManager_DeleteBlankRows_RemovesEmptyRows()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { "Data1", 100 },
                { null,    null },
                { "Data2", 200 },
                { "",      "  " }
            };

            var options = new BlankRowDeleteOptions { CheckEntireRow = true };
            var result = BlankRowManager.DeleteBlankRows(matrix, options, out int deleted);

            Assert.AreEqual(2, deleted);
            Assert.AreEqual(2, result.GetLength(0));
            Assert.AreEqual("Data1", result[0, 0]);
            Assert.AreEqual("Data2", result[1, 0]);
        }

        [TestMethod]
        public void BlankRowManager_DeleteBlankColumns_RemovesEmptyColumns()
        {
            object?[,] matrix = new object?[2, 3]
            {
                { "Val1", null, "Val3" },
                { "Val2", "",   "Val4" }
            };

            var options = new BlankRowDeleteOptions { DeleteBlankColumns = true };
            var result = BlankRowManager.DeleteBlankRows(matrix, options, out int deleted);

            Assert.AreEqual(1, deleted);
            Assert.AreEqual(2, result.GetLength(1));
            Assert.AreEqual("Val1", result[0, 0]);
            Assert.AreEqual("Val3", result[0, 1]);
        }

        #endregion

        #region Batch 4: Find, Select, and Compare Tests

        [TestMethod]
        public void DuplicateAnalyzer_DuplicatesExcludingFirst_FindsLaterOccurrences()
        {
            // 3x3 matrix:
            // "A", "B", "C"
            // "a", "D", "B"
            // "E", "A", "F"
            object?[,] matrix = new object?[3, 3]
            {
                { "A", "B", "C" },
                { "a", "D", "B" },
                { "E", "A", "F" }
            };

            var options = new DuplicateAnalysisOptions
            {
                MatchRule = DuplicateMatchRule.DuplicatesExcludingFirst,
                Scope = DuplicateAnalysisScope.CellByCell,
                MatchCase = false
            };

            var result = DuplicateAnalyzer.Analyze(matrix, 1, 1, options);

            // "a" at (2,1), "B" at (2,3), "A" at (3,2) should be reported
            Assert.AreEqual(3, result.MatchingCells.Count);
            Assert.AreEqual(9, result.TotalCellsAnalyzed);
            Assert.AreEqual(2, result.DuplicateGroupCount); // Group "a" and Group "b"
            Assert.AreEqual(4, result.UniqueValueCount);    // "C", "D", "E", "F"
        }

        [TestMethod]
        public void DuplicateAnalyzer_AllDuplicates_FindsAllOccurringMoreThanOnce()
        {
            object?[,] matrix = new object?[3, 3]
            {
                { "A", "B", "C" },
                { "a", "D", "B" },
                { "E", "A", "F" }
            };

            var options = new DuplicateAnalysisOptions
            {
                MatchRule = DuplicateMatchRule.AllDuplicates,
                Scope = DuplicateAnalysisScope.CellByCell,
                MatchCase = false
            };

            var result = DuplicateAnalyzer.Analyze(matrix, 1, 1, options);

            // All occurrences of "A" (3 cells) and "B" (2 cells) = 5 cells
            Assert.AreEqual(5, result.MatchingCells.Count);
        }

        [TestMethod]
        public void DuplicateAnalyzer_UniqueValuesOnly_FindsSingletons()
        {
            object?[,] matrix = new object?[3, 3]
            {
                { "A", "B", "C" },
                { "a", "D", "B" },
                { "E", "A", "F" }
            };

            var options = new DuplicateAnalysisOptions
            {
                MatchRule = DuplicateMatchRule.UniqueValuesOnly,
                Scope = DuplicateAnalysisScope.CellByCell,
                MatchCase = false
            };

            var result = DuplicateAnalyzer.Analyze(matrix, 1, 1, options);

            // "C", "D", "E", "F" = 4 unique cells
            Assert.AreEqual(4, result.MatchingCells.Count);
        }

        [TestMethod]
        public void DuplicateAnalyzer_EntireRows_IdentifiesDuplicateRecords()
        {
            object?[,] matrix = new object?[4, 3]
            {
                { "John", "Doe", "Engineering" },
                { "Jane", "Smith", "Design" },
                { "John", "Doe", "Engineering" }, // Duplicate of row 0
                { "Bob",  "Brown", "Sales" }
            };

            var options = new DuplicateAnalysisOptions
            {
                MatchRule = DuplicateMatchRule.DuplicatesExcludingFirst,
                Scope = DuplicateAnalysisScope.EntireRows,
                HasHeaders = false
            };

            var result = DuplicateAnalyzer.Analyze(matrix, 1, 1, options);

            // Row 2 is duplicate -> 3 cells in row 2 (sheet row 3)
            Assert.AreEqual(3, result.MatchingCells.Count);
            Assert.IsTrue(result.MatchingCells.TrueForAll(c => c.Row == 3));
        }

        [TestMethod]
        public void MaxMinSelector_EntireSelection_FindsMaxAndMin()
        {
            object?[,] matrix = new object?[2, 3]
            {
                { 10,  42, 100 },
                { -5,  20,  99 }
            };

            // Max count 1
            var maxOptions = new MaxMinSelectionOptions
            {
                Target = MaxMinTarget.Maximum,
                Scope = MaxMinScope.EntireSelection,
                Count = 1
            };
            var maxResults = MaxMinSelector.FindExtremeCells(matrix, 1, 1, maxOptions);
            Assert.AreEqual(1, maxResults.Count);
            Assert.AreEqual(100.0, maxResults[0].Value);
            Assert.AreEqual(1, maxResults[0].Row);
            Assert.AreEqual(3, maxResults[0].Column);

            // Min count 2
            var minOptions = new MaxMinSelectionOptions
            {
                Target = MaxMinTarget.Minimum,
                Scope = MaxMinScope.EntireSelection,
                Count = 2
            };
            var minResults = MaxMinSelector.FindExtremeCells(matrix, 1, 1, minOptions);
            Assert.AreEqual(2, minResults.Count);
            Assert.AreEqual(-5.0, minResults[0].Value);
            Assert.AreEqual(10.0, minResults[1].Value);
        }

        [TestMethod]
        public void MaxMinSelector_EachRow_FindsMaxPerRow()
        {
            object?[,] matrix = new object?[2, 3]
            {
                { 10,  42, 100 },
                { -5,  20,  99 }
            };

            var options = new MaxMinSelectionOptions
            {
                Target = MaxMinTarget.Maximum,
                Scope = MaxMinScope.EachRow,
                Count = 1
            };
            var results = MaxMinSelector.FindExtremeCells(matrix, 1, 1, options);
            Assert.AreEqual(2, results.Count);
            Assert.AreEqual(100.0, results[0].Value); // Row 1
            Assert.AreEqual(99.0, results[1].Value);  // Row 2
        }

        [TestMethod]
        public void MaxMinSelector_EachColumn_FindsMinPerColumn()
        {
            object?[,] matrix = new object?[2, 3]
            {
                { 10,  42, 100 },
                { -5,  20,  99 }
            };

            var options = new MaxMinSelectionOptions
            {
                Target = MaxMinTarget.Minimum,
                Scope = MaxMinScope.EachColumn,
                Count = 1
            };
            var results = MaxMinSelector.FindExtremeCells(matrix, 1, 1, options);
            Assert.AreEqual(3, results.Count);
            Assert.AreEqual(-5.0, results[0].Value); // Col 1
            Assert.AreEqual(20.0, results[1].Value); // Col 2
            Assert.AreEqual(99.0, results[2].Value); // Col 3
        }

        [TestMethod]
        public void RowDifferenceFinder_FindsRowDifferencesFromBaseline()
        {
            object?[,] matrix = new object?[3, 4]
            {
                { "X",  "X",  "Y",  "X" },  // Col 2 differs from Col 0
                { "10", "10", "10", "20" }, // Col 3 differs from Col 0
                { "A",  "A",  "A",  "A" }   // No differences
            };

            var options = new DifferenceOptions
            {
                Direction = DifferenceDirection.RowDifferences,
                BaselineIndex = 0
            };

            var diffs = RowDifferenceFinder.FindDifferences(matrix, 1, 1, options);
            Assert.AreEqual(2, diffs.Count);
            Assert.AreEqual("Y", diffs[0].Value?.ToString());
            Assert.AreEqual("20", diffs[1].Value?.ToString());
        }

        [TestMethod]
        public void RowDifferenceFinder_FindsColumnDifferencesFromBaseline()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { "Header1", "Header2" },
                { "Header1", "Other" },   // Col 1 differs from Row 0
                { "Diff",    "Header2" }, // Col 0 differs from Row 0
                { "Header1", "Header2" }
            };

            var options = new DifferenceOptions
            {
                Direction = DifferenceDirection.ColumnDifferences,
                BaselineIndex = 0
            };

            var diffs = RowDifferenceFinder.FindDifferences(matrix, 1, 1, options);
            Assert.AreEqual(2, diffs.Count);
            Assert.IsTrue(diffs.Exists(d => d.Value?.ToString() == "Diff"));
            Assert.IsTrue(diffs.Exists(d => d.Value?.ToString() == "Other"));
        }

        [TestMethod]
        public void MultiSheetFindReplace_FindMatches_TextAndRegex()
        {
            object?[,] matrix = new object?[2, 2]
            {
                { "The quick brown fox", "fox jumps over dog" },
                { "No match here",       "Fox in the forest" }
            };

            var textOptions = new FindReplaceOptions
            {
                SearchText = "fox",
                ReplaceText = "cat",
                MatchCase = false
            };

            var textMatches = MultiSheetFindReplace.FindMatches("Sheet1", 1, 1, matrix, textOptions);
            Assert.AreEqual(3, textMatches.Count);
            Assert.AreEqual("The quick brown cat", textMatches[0].ReplacedText);

            var regexOptions = new FindReplaceOptions
            {
                SearchText = @"^No\s\w+",
                ReplaceText = "Match found",
                UseRegex = true
            };

            var regexMatches = MultiSheetFindReplace.FindMatches("Sheet1", 1, 1, matrix, regexOptions);
            Assert.AreEqual(1, regexMatches.Count);
            Assert.AreEqual("Match found here", regexMatches[0].ReplacedText);
        }

        [TestMethod]
        public void MultiSheetFindReplace_ReplaceRange_InPlace()
        {
            object?[,] matrix = new object?[2, 2]
            {
                { "Item 101", "Item 102" },
                { "Static",   "Item 103" }
            };

            var options = new FindReplaceOptions
            {
                SearchText = "Item",
                ReplaceText = "SKU",
                MatchCase = true
            };

            var replaced = MultiSheetFindReplace.ReplaceInMatrix(matrix, options, out int count);
            Assert.AreEqual(3, count);
            Assert.AreEqual("SKU 101", replaced[0, 0]);
            Assert.AreEqual("SKU 102", replaced[0, 1]);
            Assert.AreEqual("Static", replaced[1, 0]);
            Assert.AreEqual("SKU 103", replaced[1, 1]);
        }

        [TestMethod]
        public void CommentFilter_FiltersByTextAndAuthor()
        {
            var comments = new List<CommentInfo>
            {
                new CommentInfo("Sheet1", "A1", 1, 1, "Alice", "TODO: review formula"),
                new CommentInfo("Sheet1", "B2", 2, 2, "Bob", "Urgent fix required"),
                new CommentInfo("Sheet2", "C3", 3, 3, "Alice", "Review completed")
            };

            // Filter by text "review"
            var textFiltered = CommentFilter.Filter(comments, new CommentFilterOptions { FilterText = "review" });
            Assert.AreEqual(2, textFiltered.Count);

            // Filter by author "Bob"
            var authorFiltered = CommentFilter.Filter(comments, new CommentFilterOptions { FilterAuthor = "Bob" });
            Assert.AreEqual(1, authorFiltered.Count);
            Assert.AreEqual("B2", authorFiltered[0].CellAddress);

            // Filter by author "Alice" AND text "completed"
            var comboFiltered = CommentFilter.Filter(comments, new CommentFilterOptions { FilterAuthor = "Alice", FilterText = "completed" });
            Assert.AreEqual(1, comboFiltered.Count);
            Assert.AreEqual("C3", comboFiltered[0].CellAddress);
        }

        #endregion

        #region Batch 5: Worksheet and Workbook Management Tests

        [TestMethod]
        public void SheetCombiner_CombinesMultipleSheets_KeepFirstHeaderOnly()
        {
            object?[,] sheet1Data = new object?[3, 2]
            {
                { "ID", "Name" },
                { 1,    "Alice" },
                { 2,    "Bob" }
            };

            object?[,] sheet2Data = new object?[2, 2]
            {
                { "ID", "Name" },
                { 3,    "Charlie" }
            };

            var sheets = new List<SheetDataInfo>
            {
                new SheetDataInfo("North", sheet1Data),
                new SheetDataInfo("South", sheet2Data)
            };

            var options = new SheetCombineOptions
            {
                HasHeaders = true,
                KeepFirstSheetHeaderOnly = true,
                AddSourceSheetColumn = true,
                BlankRowsBetweenSheets = 0
            };

            var combined = SheetCombiner.Combine(sheets, options, out int totalRows);

            Assert.AreEqual(4, totalRows);
            Assert.AreEqual(4, combined.GetLength(0));
            Assert.AreEqual(3, combined.GetLength(1)); // Source Col + ID + Name

            // Header row
            Assert.AreEqual("Source Sheet", combined[0, 0]);
            Assert.AreEqual("ID", combined[0, 1]);
            Assert.AreEqual("Name", combined[0, 2]);

            // Sheet 1 data
            Assert.AreEqual("North", combined[1, 0]);
            Assert.AreEqual(1, combined[1, 1]);
            Assert.AreEqual("Alice", combined[1, 2]);

            Assert.AreEqual("North", combined[2, 0]);
            Assert.AreEqual(2, combined[2, 1]);
            Assert.AreEqual("Bob", combined[2, 2]);

            // Sheet 2 data (header omitted, row appended)
            Assert.AreEqual("South", combined[3, 0]);
            Assert.AreEqual(3, combined[3, 1]);
            Assert.AreEqual("Charlie", combined[3, 2]);
        }

        [TestMethod]
        public void SheetCombiner_BlankRowsSpacing_InsertsConfiguredBlanks()
        {
            object?[,] s1 = new object?[2, 1] { { "H" }, { "A" } };
            object?[,] s2 = new object?[2, 1] { { "H" }, { "B" } };

            var sheets = new List<SheetDataInfo>
            {
                new SheetDataInfo("S1", s1),
                new SheetDataInfo("S2", s2)
            };

            var options = new SheetCombineOptions
            {
                HasHeaders = true,
                KeepFirstSheetHeaderOnly = true,
                AddSourceSheetColumn = false,
                BlankRowsBetweenSheets = 1
            };

            var combined = SheetCombiner.Combine(sheets, options, out int totalRows);

            // Row 0: H, Row 1: A, Row 2: (blank), Row 3: B -> 4 rows
            Assert.AreEqual(4, totalRows);
            Assert.AreEqual("H", combined[0, 0]);
            Assert.AreEqual("A", combined[1, 0]);
            Assert.IsNull(combined[2, 0]);
            Assert.AreEqual("B", combined[3, 0]);
        }

        [TestMethod]
        public void SheetTabSorter_NaturalNumeric_CorrectlyOrdersEmbeddedNumbers()
        {
            var raw = new List<string> { "Sheet10", "Sheet2", "Sheet1", "Sheet20" };
            var sorted = SheetTabSorter.Sort(raw, SheetSortMode.NaturalNumeric);

            Assert.AreEqual("Sheet1", sorted[0]);
            Assert.AreEqual("Sheet2", sorted[1]);
            Assert.AreEqual("Sheet10", sorted[2]);
            Assert.AreEqual("Sheet20", sorted[3]);
        }

        [TestMethod]
        public void SheetTabSorter_AlphabeticalAndReverse_SortsCorrectly()
        {
            var raw = new List<string> { "Zebra", "Apple", "Mango" };

            var asc = SheetTabSorter.Sort(raw, SheetSortMode.AlphabeticalAscending);
            Assert.AreEqual("Apple", asc[0]);
            Assert.AreEqual("Mango", asc[1]);
            Assert.AreEqual("Zebra", asc[2]);

            var desc = SheetTabSorter.Sort(raw, SheetSortMode.AlphabeticalDescending);
            Assert.AreEqual("Zebra", desc[0]);
            Assert.AreEqual("Mango", desc[1]);
            Assert.AreEqual("Apple", desc[2]);

            var rev = SheetTabSorter.Sort(raw, SheetSortMode.ReverseCurrent);
            Assert.AreEqual("Mango", rev[0]);
            Assert.AreEqual("Apple", rev[1]);
            Assert.AreEqual("Zebra", rev[2]);
        }

        [TestMethod]
        public void SheetIndexGenerator_GeneratesEntriesAndFormulas()
        {
            var sheets = new List<string> { "Summary", "Q1 Sales", "Settings" };
            var entries = SheetIndexGenerator.GenerateEntries(sheets, "Index");

            Assert.AreEqual(3, entries.Count);
            Assert.AreEqual(1, entries[0].SequenceNumber);
            Assert.AreEqual("Summary", entries[0].SheetName);
            Assert.AreEqual("=HYPERLINK(\"#'Summary'!A1\", \"Go to Summary\")", entries[0].HyperlinkFormula);

            // Back-link formula
            string backLink = SheetIndexGenerator.CreateBackLinkFormula("Index");
            Assert.AreEqual("=HYPERLINK(\"#'Index'!A1\", \"[Back to Index]\")", backLink);
        }

        [TestMethod]
        public void BlankSheetDetector_IdentifiesBlankMatrix()
        {
            // Null matrix is blank
            Assert.IsTrue(BlankSheetDetector.IsMatrixBlank(null));

            // Matrix containing only nulls and empty whitespace is blank
            object?[,] emptyMatrix = new object?[2, 2]
            {
                { null, "" },
                { "   ", "\t" }
            };
            Assert.IsTrue(BlankSheetDetector.IsMatrixBlank(emptyMatrix));

            // Matrix containing data is NOT blank
            object?[,] dataMatrix = new object?[2, 2]
            {
                { null, null },
                { null, "Real Content" }
            };
            Assert.IsFalse(BlankSheetDetector.IsMatrixBlank(dataMatrix));
        }

        [TestMethod]
        public void WorkbookSplitter_ResolvesPatternAndSanitizes()
        {
            string sanitized = WorkbookSplitter.SanitizeFileName("Quarter 1/2*Data?");
            Assert.AreEqual("Quarter 1_2_Data", sanitized);

            string resolved = WorkbookSplitter.ResolveFileName(
                "[WorkbookName]_[SheetName]",
                "AnnualReport.xlsx",
                "Q1:Sales",
                "pdf");

            Assert.AreEqual("AnnualReport_Q1_Sales.pdf", resolved);
        }

        #endregion

        #region Batch 6: Data, Filtering, and Validation Tests

        [TestMethod]
        public void AdvancedSorter_SortsByTextLength_AscendingAndDescending()
        {
            object?[,] matrix = new object?[5, 1]
            {
                { "Header" },
                { "Watermelon" }, // 10
                { "Fig" },        // 3
                { "Banana" },     // 6
                { "Pineapple" }   // 9
            };

            // Ascending
            var ascOptions = new AdvancedSortOptions
            {
                SortColumnIndex = 0,
                HasHeaderRow = true,
                Criterion = AdvancedSortCriterion.TextLengthAscending
            };
            var ascResult = AdvancedSorter.Sort(matrix, ascOptions);

            Assert.AreEqual("Header", ascResult[0, 0]);
            Assert.AreEqual("Fig", ascResult[1, 0]);
            Assert.AreEqual("Banana", ascResult[2, 0]);
            Assert.AreEqual("Pineapple", ascResult[3, 0]);
            Assert.AreEqual("Watermelon", ascResult[4, 0]);

            // Descending
            var descOptions = new AdvancedSortOptions
            {
                SortColumnIndex = 0,
                HasHeaderRow = true,
                Criterion = AdvancedSortCriterion.TextLengthDescending
            };
            var descResult = AdvancedSorter.Sort(matrix, descOptions);

            Assert.AreEqual("Header", descResult[0, 0]);
            Assert.AreEqual("Watermelon", descResult[1, 0]);
            Assert.AreEqual("Pineapple", descResult[2, 0]);
            Assert.AreEqual("Banana", descResult[3, 0]);
            Assert.AreEqual("Fig", descResult[4, 0]);
        }

        [TestMethod]
        public void AdvancedSorter_SortsByMonthName()
        {
            object?[,] matrix = new object?[4, 1]
            {
                { "March" },
                { "December" },
                { "January" },
                { "August" }
            };

            var options = new AdvancedSortOptions
            {
                SortColumnIndex = 0,
                HasHeaderRow = false,
                Criterion = AdvancedSortCriterion.MonthName
            };
            var result = AdvancedSorter.Sort(matrix, options);

            Assert.AreEqual("January", result[0, 0]);
            Assert.AreEqual("March", result[1, 0]);
            Assert.AreEqual("August", result[2, 0]);
            Assert.AreEqual("December", result[3, 0]);
        }

        [TestMethod]
        public void AdvancedSorter_SortsByWeekdayName()
        {
            object?[,] matrix = new object?[4, 1]
            {
                { "Friday" },
                { "Monday" },
                { "Sunday" },
                { "Wednesday" }
            };

            var options = new AdvancedSortOptions
            {
                SortColumnIndex = 0,
                HasHeaderRow = false,
                Criterion = AdvancedSortCriterion.WeekdayName
            };
            var result = AdvancedSorter.Sort(matrix, options);

            Assert.AreEqual("Sunday", result[0, 0]);
            Assert.AreEqual("Monday", result[1, 0]);
            Assert.AreEqual("Wednesday", result[2, 0]);
            Assert.AreEqual("Friday", result[3, 0]);
        }

        [TestMethod]
        public void ColorAnalyzer_FiltersCellsByColor()
        {
            var cells = new List<CellColorInfo>
            {
                new CellColorInfo(1, 1, 0xFF0000, 0x000000, 100),
                new CellColorInfo(1, 2, 0x00FF00, 0x000000, 200),
                new CellColorInfo(2, 1, 0xFF0000, 0x000000, 300)
            };

            var redMatches = ColorAnalyzer.FilterCellsByColor(cells, 0xFF0000, isFillColor: true);
            Assert.AreEqual(2, redMatches.Count);
            Assert.IsTrue(redMatches.Contains((1, 1)));
            Assert.IsTrue(redMatches.Contains((2, 1)));

            var greenMatches = ColorAnalyzer.FilterCellsByColor(cells, 0x00FF00, isFillColor: true);
            Assert.AreEqual(1, greenMatches.Count);
            Assert.IsTrue(greenMatches.Contains((1, 2)));
        }

        [TestMethod]
        public void ColorAnalyzer_CalculatesColorStatistics()
        {
            var cells = new List<CellColorInfo>
            {
                new CellColorInfo(1, 1, 0xFF0000, 0x000000, 100),
                new CellColorInfo(1, 2, 0x00FF00, 0x000000, 200),
                new CellColorInfo(2, 1, 0xFF0000, 0x000000, 300)
            };

            var stats = ColorAnalyzer.CalculateColorStatistics(cells, isFillColor: true);
            Assert.AreEqual(2, stats.Count);

            var redGroup = stats.Find(s => s.ColorRgb == 0xFF0000);
            Assert.IsNotNull(redGroup);
            Assert.AreEqual(2, redGroup.CellCount);
            Assert.AreEqual(400.0, redGroup.Sum);
            Assert.AreEqual(200.0, redGroup.Average);
            Assert.AreEqual(100.0, redGroup.Min);
            Assert.AreEqual(300.0, redGroup.Max);

            var greenGroup = stats.Find(s => s.ColorRgb == 0x00FF00);
            Assert.IsNotNull(greenGroup);
            Assert.AreEqual(1, greenGroup.CellCount);
            Assert.AreEqual(200.0, greenGroup.Sum);
        }

        [TestMethod]
        public void TableMerger_JoinsOnKey_UpdatesColumns()
        {
            object?[,] primary = new object?[3, 2]
            {
                { "ID",  "Name" },
                { 101,   "Alice" },
                { 102,   "Bob" }
            };

            object?[,] lookup = new object?[4, 2]
            {
                { "ID",  "Department" },
                { 101,   "HR" },
                { 102,   "Finance" },
                { 103,   "Engineering" }
            };

            var options = new TableMergeOptions
            {
                PrimaryKeyColumn = 0,
                LookupKeyColumn = 0,
                HasHeaders = true,
                AppendUnmatchedRows = false
            };

            var merged = TableMerger.Merge(primary, lookup, options, out int matched, out int unmatched);

            Assert.AreEqual(2, matched);
            Assert.AreEqual(0, unmatched);
            Assert.AreEqual(3, merged.GetLength(0)); // Header + 2 rows
            Assert.AreEqual(3, merged.GetLength(1)); // ID, Name, Department

            // Check header
            Assert.AreEqual("Department", merged[0, 2]);

            // Check matched data
            Assert.AreEqual("HR", merged[1, 2]);
            Assert.AreEqual("Finance", merged[2, 2]);
        }

        [TestMethod]
        public void TableMerger_AppendUnmatchedRows_IncludesRemainingLookupRows()
        {
            object?[,] primary = new object?[2, 2]
            {
                { "ID",  "Name" },
                { 101,   "Alice" }
            };

            object?[,] lookup = new object?[3, 2]
            {
                { "ID",  "Department" },
                { 101,   "HR" },
                { 102,   "Finance" }
            };

            var options = new TableMergeOptions
            {
                PrimaryKeyColumn = 0,
                LookupKeyColumn = 0,
                HasHeaders = true,
                AppendUnmatchedRows = true
            };

            var merged = TableMerger.Merge(primary, lookup, options, out int matched, out int unmatched);

            Assert.AreEqual(1, matched);
            Assert.AreEqual(3, merged.GetLength(0)); // Header, row 101, appended row 102
            Assert.AreEqual(102, merged[2, 0]);
            Assert.AreEqual("Finance", merged[2, 2]);
        }

        [TestMethod]
        public void DistinctExtractor_ExtractsUniqueRowsOnly()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { "Category", "Item" },
                { "Fruit",    "Apple" },
                { "Fruit",    "Apple" }, // duplicate
                { "Produce",  "Carrot" }
            };

            var options = new DistinctExtractOptions
            {
                HasHeaders = true,
                KeyColumns = null, // entire row
                MatchCase = false
            };

            var distinct = DistinctExtractor.ExtractDistinct(matrix, options, out int unique, out int duplicate);

            Assert.AreEqual(2, unique);
            Assert.AreEqual(1, duplicate);
            Assert.AreEqual(3, distinct.GetLength(0)); // Header + 2 unique rows
            Assert.AreEqual("Apple", distinct[1, 1]);
            Assert.AreEqual("Carrot", distinct[2, 1]);
        }

        [TestMethod]
        public void DuplicateValidationHelper_BuildsFormula()
        {
            string formula = DuplicateValidationHelper.BuildCountIfFormula("C", 5);
            Assert.AreEqual("=COUNTIF(C:C, C5)<=1", formula);

            string colLetter = DuplicateValidationHelper.IndexToColumnLetter(28);
            Assert.AreEqual("AB", colLetter);
        }

        #region Batch 7: Formulas, Math Operations & Page Setup Tests

        [TestMethod]
        public void FormulaReferenceConverter_ToAbsolute_ConvertsCorrectly()
        {
            string formula = "=SUM(A1:B10) + C5 * $D$6";
            string converted = FormulaReferenceConverter.ConvertFormulaReferences(formula, ReferenceConversionMode.ToAbsolute);
            Assert.AreEqual("=SUM($A$1:$B$10) + $C$5 * $D$6", converted);

            // Preserves sheet qualifier and string literal
            string complex = "='Sales Data'!A1 + Sheet2!B2 & \" Address: A1 \"";
            string resComplex = FormulaReferenceConverter.ConvertFormulaReferences(complex, ReferenceConversionMode.ToAbsolute);
            Assert.AreEqual("='Sales Data'!$A$1 + Sheet2!$B$2 & \" Address: A1 \"", resComplex);
        }

        [TestMethod]
        public void FormulaReferenceConverter_ToRowAbsolute_ConvertsCorrectly()
        {
            string formula = "=$A$1 + $B2 + C$3";
            string converted = FormulaReferenceConverter.ConvertFormulaReferences(formula, ReferenceConversionMode.ToRowAbsolute);
            Assert.AreEqual("=A$1 + B$2 + C$3", converted);
        }

        [TestMethod]
        public void FormulaReferenceConverter_ToColumnAbsolute_ConvertsCorrectly()
        {
            string formula = "=A1 + B$2 + C3";
            string converted = FormulaReferenceConverter.ConvertFormulaReferences(formula, ReferenceConversionMode.ToColumnAbsolute);
            Assert.AreEqual("=$A1 + $B2 + $C3", converted);
        }

        [TestMethod]
        public void FormulaReferenceConverter_ToRelative_ConvertsCorrectly()
        {
            string formula = "=SUM($A$1:$B$10) + $C$5 * D$6";
            string converted = FormulaReferenceConverter.ConvertFormulaReferences(formula, ReferenceConversionMode.ToRelative);
            Assert.AreEqual("=SUM(A1:B10) + C5 * D6", converted);
        }

        [TestMethod]
        public void FormulaReferenceConverter_ConvertMatrix_BatchProcessesFormulas()
        {
            object?[,] matrix = new object?[2, 2]
            {
                { "=A1+B1", "Static Text" },
                { 12345,    "=SUM(C1:D5)" }
            };

            var (res, modified) = FormulaReferenceConverter.ConvertMatrix(matrix, ReferenceConversionMode.ToAbsolute);
            Assert.AreEqual(2, modified);
            Assert.AreEqual("=$A$1+$B$1", res[0, 0]);
            Assert.AreEqual("Static Text", res[0, 1]);
            Assert.AreEqual(12345, res[1, 0]);
            Assert.AreEqual("=SUM($C$1:$D$5)", res[1, 1]);
        }

        [TestMethod]
        public void FormulaExactCopier_FillExactFormula_FillsVerbatim()
        {
            var result = FormulaExactCopier.FillExactFormula("=A1+B1", 2, 3);
            Assert.AreEqual(2, result.GetLength(0));
            Assert.AreEqual(3, result.GetLength(1));
            Assert.AreEqual("=A1+B1", result[0, 0]);
            Assert.AreEqual("=A1+B1", result[1, 2]);
        }

        [TestMethod]
        public void FormulaExactCopier_CopyExactMatrix_PreservesAllReferences()
        {
            object?[,] source = new object?[2, 1]
            {
                { "=A1*2" },
                { "=B1*2" }
            };

            var target = FormulaExactCopier.CopyExactMatrix(source, 2, 1);
            Assert.AreEqual("=A1*2", target[0, 0]);
            Assert.AreEqual("=B1*2", target[1, 0]);
        }

        [TestMethod]
        public void FormulaErrorLocator_DetectsErrors()
        {
            object?[,] values = new object?[3, 2]
            {
                { 100, FormulaErrorLocator.CVErrDiv0 },
                { "#VALUE!", "Valid Text" },
                { FormulaErrorLocator.CVErrNA, "#REF!" }
            };

            object?[,] formulas = new object?[3, 2]
            {
                { null, "=10/0" },
                { "=A1+B1", null },
                { "=VLOOKUP(1,A:B,2,FALSE)", "=#REF!*2" }
            };

            var errors = FormulaErrorLocator.LocateErrors(values, formulas, 1, 1, "Sheet1");
            Assert.AreEqual(4, errors.Count);
            Assert.AreEqual("#DIV/0!", errors[0].ErrorText);
            Assert.AreEqual("$B$1", errors[0].CellAddress);
            Assert.AreEqual("=10/0", errors[0].Formula);
            Assert.AreEqual("#VALUE!", errors[1].ErrorText);
            Assert.AreEqual("#N/A", errors[2].ErrorText);
            Assert.AreEqual("#REF!", errors[3].ErrorText);
        }

        [TestMethod]
        public void RangeMathOperations_AppliesCalculationsAccurately()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { 100.0, 50.0 },
                { "200", null },
                { "Text", 0.0 },
                { -10.0, 25.5 }
            };

            // Test Addition
            var addRes = RangeMathOperations.Apply(matrix, MathOperationType.Add, 15.0);
            Assert.AreEqual(115.0, addRes.Matrix[0, 0]);
            Assert.AreEqual(65.0, addRes.Matrix[0, 1]);
            Assert.AreEqual(215.0, addRes.Matrix[1, 0]); // parsed string
            Assert.IsNull(addRes.Matrix[1, 1]); // skipped blank
            Assert.AreEqual("Text", addRes.Matrix[2, 0]); // skipped text
            Assert.AreEqual(5.0, addRes.Matrix[3, 0]);

            // Test Multiplication with rounding
            var opts = new MathOperationOptions { RoundResults = true, DecimalPlaces = 1 };
            var mulRes = RangeMathOperations.Apply(matrix, MathOperationType.Multiply, 1.333, opts);
            Assert.AreEqual(133.3, mulRes.Matrix[0, 0]);

            // Test Division by Zero safeguard
            var divZero = RangeMathOperations.Apply(matrix, MathOperationType.Divide, 0.0);
            Assert.AreEqual("#DIV/0!", divZero.Matrix[0, 0]);
            Assert.IsTrue(divZero.ErrorCount > 0);
        }

        [TestMethod]
        public void MockExcelContext_PageSetupAndFreezePanes()
        {
            var ctx = new MockExcelContext();
            var options = new PageSetupOptions
            {
                CopyMargins = true,
                CopyOrientationAndSize = true
            };

            int copied = ctx.CopyPageSetup("Sheet1", new[] { "Sheet2", "Sheet3" }, options);
            Assert.AreEqual(2, copied);
            Assert.AreEqual(1, ctx.CopiedPageSetups.Count);
            Assert.AreEqual("Sheet1", ctx.CopiedPageSetups[0].src);
            Assert.AreEqual(2, ctx.CopiedPageSetups[0].targets.Count);

            int frozen = ctx.SynchronizeFreezePanes(new[] { "Sheet1", "Sheet2" }, FreezePanesSyncMode.FreezeTopRow, 2, 1);
            Assert.AreEqual(2, frozen);
            Assert.IsNotNull(ctx.LastSynchronizedFreezePanes);
            Assert.AreEqual(FreezePanesSyncMode.FreezeTopRow, ctx.LastSynchronizedFreezePanes.Value.mode);
        }

        #region Batch 8: Import, Export & File Utilities Tests

        [TestMethod]
        public void CsvTextImporter_ParseText_ParsesStandardCsv()
        {
            string csv = "Name,Age,Score\r\nAlice,30,95.5\r\nBob,25,88.0";
            var matrix = CsvTextImporter.ParseText(csv);

            Assert.AreEqual(3, matrix.GetLength(0));
            Assert.AreEqual(3, matrix.GetLength(1));
            Assert.AreEqual("Name", matrix[0, 0]);
            Assert.AreEqual("Alice", matrix[1, 0]);
            Assert.AreEqual(30.0, matrix[1, 1]);
            Assert.AreEqual(95.5, matrix[1, 2]);
        }

        [TestMethod]
        public void CsvTextImporter_ParseText_HandlesQuotedFieldsAndEscapes()
        {
            string csv = "\"Item\",\"Description\",\"Price\"\r\n\"Widget A\",\"Contains \"\"special\"\" parts, and commas\",19.99\r\n\"Widget B\",\"Line 1\r\nLine 2\",29.5";
            var matrix = CsvTextImporter.ParseText(csv);

            Assert.AreEqual(3, matrix.GetLength(0));
            Assert.AreEqual(3, matrix.GetLength(1));
            Assert.AreEqual("Contains \"special\" parts, and commas", matrix[1, 1]);
            Assert.AreEqual("Line 1\r\nLine 2", matrix[2, 1]);
            Assert.AreEqual(19.99, matrix[1, 2]);
        }

        [TestMethod]
        public void CsvTextImporter_ParseText_HandlesCustomDelimiters()
        {
            string pipeText = "SKU|Product|Qty\r\n1001|Keyboard|15\r\n1002|Mouse|30";
            var options = new CsvImportOptions { Delimiter = "|" };
            var matrix = CsvTextImporter.ParseText(pipeText, options);

            Assert.AreEqual(3, matrix.GetLength(0));
            Assert.AreEqual(3, matrix.GetLength(1));
            Assert.AreEqual("Keyboard", matrix[1, 1]);
            Assert.AreEqual(15.0, matrix[1, 2]);
        }

        [TestMethod]
        public void RangeExporter_ExportToString_ExportsCsvAndTsv()
        {
            object?[,] matrix = new object?[2, 2]
            {
                { "Header1", "Header, 2" },
                { "Hello \"World\"", 123.45 }
            };

            string csv = RangeExporter.ExportToString(matrix, new RangeExportOptions { Format = ExportFileFormat.Csv });
            Assert.IsTrue(csv.Contains("\"Header, 2\""));
            Assert.IsTrue(csv.Contains("\"Hello \"\"World\"\"\""));

            string tsv = RangeExporter.ExportToString(matrix, new RangeExportOptions { Format = ExportFileFormat.Tsv });
            Assert.IsTrue(tsv.Contains("Header1\t"));
        }

        [TestMethod]
        public void RangeExporter_ExportToString_ExportsJson()
        {
            object?[,] matrix = new object?[3, 2]
            {
                { "Name", "Score" },
                { "Alice", 95 },
                { "Bob", 88.5 }
            };

            string json = RangeExporter.ExportToString(matrix, new RangeExportOptions { Format = ExportFileFormat.Json, HasHeaders = true });
            Assert.IsTrue(json.Contains("\"Name\": \"Alice\""));
            Assert.IsTrue(json.Contains("\"Score\": 95"));
            Assert.IsTrue(json.Contains("\"Name\": \"Bob\""));
        }

        [TestMethod]
        public void FolderInventoryScanner_ScanFolder_ScansAndBuildsTable()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "EPSTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try
            {
                File.WriteAllText(Path.Combine(tempDir, "test1.txt"), "Sample content 1");
                File.WriteAllText(Path.Combine(tempDir, "test2.csv"), "a,b,c");

                var res = FolderInventoryScanner.ScanFolder(tempDir);
                Assert.AreEqual(2, res.FileCount);
                Assert.AreEqual(3, res.TableData.GetLength(0)); // Header + 2 files
                Assert.AreEqual("File Name", res.TableData[0, 0]);
                Assert.AreEqual("Link", res.TableData[0, 7]);
                Assert.IsTrue(res.TableData[1, 7]?.ToString()?.StartsWith("=HYPERLINK"));
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [TestMethod]
        public void TextFileReader_ReadFileToMatrix_ReadsRowsAndSingleCell()
        {
            string tempFile = Path.GetTempFileName();
            try
            {
                File.WriteAllLines(tempFile, new[] { "Line 1", "Line 2", "Line 3" });

                var rowByRow = TextFileReader.ReadFileToMatrix(tempFile, TextFileInsertMode.RowByRow);
                Assert.AreEqual(3, rowByRow.GetLength(0));
                Assert.AreEqual("Line 2", rowByRow[1, 0]);

                var singleCell = TextFileReader.ReadFileToMatrix(tempFile, TextFileInsertMode.SingleCell);
                Assert.AreEqual(1, singleCell.GetLength(0));
                Assert.AreEqual(1, singleCell.GetLength(1));
                Assert.IsTrue(singleCell[0, 0]?.ToString()?.Contains("Line 3"));
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [TestMethod]
        public void MockExcelContext_ExportPdfAndBatchConvert()
        {
            var ctx = new MockExcelContext();
            bool exported = ctx.ExportToPdf("C:\\output.pdf", new[] { "Sheet1" }, true);
            Assert.IsTrue(exported);
            Assert.AreEqual(1, ctx.ExportedPdfs.Count);
            Assert.AreEqual("C:\\output.pdf", ctx.ExportedPdfs[0].dest);
            Assert.IsTrue(ctx.ExportedPdfs[0].open);

            int converted = ctx.BatchConvertWorkbooks(new[] { "C:\\f1.xls", "C:\\f2.xls" }, "C:\\Out", ".xlsx");
            Assert.AreEqual(2, converted);
            Assert.AreEqual(1, ctx.ConvertedWorkbooks.Count);
            Assert.AreEqual(".xlsx", ctx.ConvertedWorkbooks[0].ext);
        }

        #endregion

        #region Batch 9: Protection, Recovery, Charts & Analysis Tests

        [TestMethod]
        public void QuickStatisticsEngine_EmptyArray_ReturnsZeroCounts()
        {
            var res = QuickStatisticsEngine.Calculate(null);
            Assert.AreEqual(0, res.TotalCells);
            Assert.AreEqual(0, res.NumericCells);
            Assert.AreEqual(0, res.BlankCells);
            Assert.AreEqual(0, res.Sum);

            var emptyRes = QuickStatisticsEngine.Calculate(new object?[0, 0]);
            Assert.AreEqual(0, emptyRes.TotalCells);
        }

        [TestMethod]
        public void QuickStatisticsEngine_MixedData_CalculatesCorrectMetrics()
        {
            object?[,] matrix = new object?[2, 3]
            {
                { 10.0, 20.0, 30.0 },
                { 40.0, null, "Text" }
            };

            var res = QuickStatisticsEngine.Calculate(matrix);

            Assert.AreEqual(6, res.TotalCells);
            Assert.AreEqual(4, res.NumericCells);
            Assert.AreEqual(1, res.BlankCells);
            Assert.AreEqual(1, res.TextCells);
            Assert.AreEqual(0, res.ErrorCells);
            Assert.AreEqual(100.0, res.Sum);
            Assert.AreEqual(25.0, res.Mean);
            Assert.AreEqual(25.0, res.Median); // (20 + 30) / 2
            Assert.AreEqual(10.0, res.Min);
            Assert.AreEqual(40.0, res.Max);
            Assert.AreEqual(30.0, res.Range);
            Assert.AreEqual(4, res.DistinctNumericCount);
        }

        [TestMethod]
        public void QuickStatisticsEngine_AdvancedDispersion_VarianceAndSkewness()
        {
            // Set of 8 numbers: { 2, 4, 4, 4, 5, 5, 7, 9 }
            // Sum = 40, Mean = 5.0, Mode = 4.0
            // Diffs from 5: -3, -1, -1, -1, 0, 0, 2, 4 -> Sq: 9, 1, 1, 1, 0, 0, 4, 16 -> SumSq = 32
            // PopVar = 32 / 8 = 4.0 -> PopStdDev = 2.0
            // SampleVar = 32 / 7 = 4.5714...
            object?[,] matrix = new object?[8, 1]
            {
                { 2 }, { 4 }, { 4 }, { 4 }, { 5 }, { 5 }, { 7 }, { 9 }
            };

            var res = QuickStatisticsEngine.Calculate(matrix);

            Assert.AreEqual(8, res.NumericCells);
            Assert.AreEqual(40.0, res.Sum);
            Assert.AreEqual(5.0, res.Mean);
            Assert.AreEqual(4.0, res.Mode);
            Assert.AreEqual(4.5, res.Median);
            Assert.AreEqual(4.0, res.PopulationVariance, 1e-4);
            Assert.AreEqual(2.0, res.PopulationStdDev, 1e-4);
            Assert.AreEqual(32.0 / 7.0, res.SampleVariance, 1e-4);
            Assert.IsTrue(res.SampleStdDev > 2.0);
            Assert.IsNotNull(res.Skewness);
            Assert.IsTrue(res.Skewness > 0); // positive skew due to 7 and 9
        }

        [TestMethod]
        public void QuickStatisticsEngine_GenerateSummaryTable_CreatesProper2DMatrix()
        {
            object?[,] matrix = new object?[3, 1]
            {
                { 100 }, { 200 }, { 300 }
            };

            var res = QuickStatisticsEngine.Calculate(matrix);
            var table = QuickStatisticsEngine.GenerateSummaryTable(res, "$A$1:$A$3");

            Assert.AreEqual("Statistic", table[0, 0]);
            Assert.AreEqual("Value", table[0, 1]);
            Assert.AreEqual("$A$1:$A$3", table[1, 1]);
            Assert.AreEqual(600.0, Convert.ToDouble(table[8, 1])); // Sum row
        }

        [TestMethod]
        public void WorkbookSnapshotManager_GenerateAndParsePaths()
        {
            string dir = Path.GetTempPath();
            string path = WorkbookSnapshotManager.GenerateSnapshotFilePath(dir, "Financial Report 2026.xlsx", "CleanData");

            Assert.IsTrue(path.Contains("Financial_Report_2026"));
            Assert.IsTrue(path.Contains("CleanData"));
            Assert.IsTrue(path.EndsWith(".xlsx"));
        }

        [TestMethod]
        public void WorkbookSnapshotManager_FileOperations_CreateListDeleteClean()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "SnapshotsTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try
            {
                string f1 = Path.Combine(tempDir, "Budget_20261028_100000_Tag1.xlsx");
                string f2 = Path.Combine(tempDir, "Budget_20261028_110000_Tag2.xlsx");
                string f3 = Path.Combine(tempDir, "Other_20261028_120000_Tag3.xlsx");

                File.WriteAllText(f1, "dummy");
                File.WriteAllText(f2, "dummy");
                File.WriteAllText(f3, "dummy");

                var allSnapshots = WorkbookSnapshotManager.GetSnapshots(tempDir);
                Assert.AreEqual(3, allSnapshots.Count);

                var budgetSnapshots = WorkbookSnapshotManager.GetSnapshots(tempDir, "Budget.xlsx");
                Assert.AreEqual(2, budgetSnapshots.Count);

                bool deleted = WorkbookSnapshotManager.DeleteSnapshot(f3);
                Assert.IsTrue(deleted);
                Assert.IsFalse(File.Exists(f3));

                int cleaned = WorkbookSnapshotManager.CleanOldSnapshots(tempDir, maxSnapshots: 1);
                Assert.AreEqual(1, cleaned);
                var remaining = WorkbookSnapshotManager.GetSnapshots(tempDir);
                Assert.AreEqual(1, remaining.Count);
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [TestMethod]
        public void TargetVsActualDetector_DetectColumnsAndSummary()
        {
            object?[,] matrix = new object?[3, 3]
            {
                { "Period", "Actual Sales", "Target Goal" },
                { "Q1",     10000.0,        8000.0 },
                { "Q2",     12000.0,        15000.0 }
            };

            var cfg = TargetVsActualDetector.Detect(matrix, 3);

            Assert.AreEqual(1, cfg.CategoryColumnIndex);
            Assert.AreEqual(2, cfg.ActualColumnIndex);
            Assert.AreEqual(3, cfg.TargetColumnIndex);
            Assert.AreEqual(2, cfg.ValidDataRowCount);
            Assert.AreEqual(22000.0, cfg.TotalActual);
            Assert.AreEqual(23000.0, cfg.TotalTarget);
            Assert.AreEqual(-1000.0, cfg.TotalVariance);
            Assert.IsTrue(cfg.VariancePercent < 0);
        }

        [TestMethod]
        public void MockExcelContext_BatchProtectAndUnprotect()
        {
            var ctx = new MockExcelContext();
            var options = new SheetProtectionOptions
            {
                FormatCells = true,
                Sort = true,
                SelectLockedCells = true
            };

            var (protCount, protErrors) = ctx.BatchProtectWorksheets(new[] { "Sheet1", "Sheet2" }, "Secret123", options);
            Assert.AreEqual(2, protCount);
            Assert.AreEqual(0, protErrors.Count);
            Assert.AreEqual(1, ctx.ProtectedSheetsHistory.Count);
            Assert.AreEqual("Secret123", ctx.ProtectedSheetsHistory[0].pwd);
            Assert.IsTrue(ctx.ProtectedSheetsHistory[0].opts.FormatCells);

            var (unprotCount, unprotErrors) = ctx.BatchUnprotectWorksheets(new[] { "Sheet1" }, "Secret123");
            Assert.AreEqual(1, unprotCount);
            Assert.AreEqual(0, unprotErrors.Count);

            bool saved = ctx.SaveActiveWorkbookCopy(@"C:\Temp\Snapshot.xlsx");
            Assert.IsTrue(saved);
            Assert.AreEqual(1, ctx.SavedCopies.Count);

            bool restored = ctx.RestoreWorkbookFromSnapshot(@"C:\Temp\Snapshot.xlsx");
            Assert.IsTrue(restored);
            Assert.AreEqual(@"C:\Temp\Snapshot.xlsx", ctx.RestoredSnapshotPath);

            bool charted = ctx.CreateTargetVsActualChart("Q1-Q4 Review", 1, 1, 5, 3, 1, 2, 3);
            Assert.IsTrue(charted);
            Assert.IsNotNull(ctx.LastCreatedChart);
            Assert.AreEqual("Q1-Q4 Review", ctx.LastCreatedChart.Value.title);
        }

        #endregion

        #region Batch 10: Navigation & User Productivity Tests

        [TestMethod]
        public void ColumnLetterConverter_IndexToLetter_ConvertsProperly()
        {
            Assert.AreEqual("A", ColumnLetterConverter.IndexToLetter(1));
            Assert.AreEqual("Z", ColumnLetterConverter.IndexToLetter(26));
            Assert.AreEqual("AA", ColumnLetterConverter.IndexToLetter(27));
            Assert.AreEqual("AB", ColumnLetterConverter.IndexToLetter(28));
            Assert.AreEqual("AZ", ColumnLetterConverter.IndexToLetter(52));
            Assert.AreEqual("BA", ColumnLetterConverter.IndexToLetter(53));
            Assert.AreEqual("ZZ", ColumnLetterConverter.IndexToLetter(702));
            Assert.AreEqual("AAA", ColumnLetterConverter.IndexToLetter(703));
        }

        [TestMethod]
        public void ColumnLetterConverter_LetterToIndex_ConvertsProperly()
        {
            Assert.AreEqual(1, ColumnLetterConverter.LetterToIndex("A"));
            Assert.AreEqual(26, ColumnLetterConverter.LetterToIndex("Z"));
            Assert.AreEqual(27, ColumnLetterConverter.LetterToIndex("AA"));
            Assert.AreEqual(28, ColumnLetterConverter.LetterToIndex("AB"));
            Assert.AreEqual(52, ColumnLetterConverter.LetterToIndex("AZ"));
            Assert.AreEqual(53, ColumnLetterConverter.LetterToIndex("BA"));
            Assert.AreEqual(702, ColumnLetterConverter.LetterToIndex("ZZ"));
            Assert.AreEqual(703, ColumnLetterConverter.LetterToIndex("AAA"));
        }

        [TestMethod]
        public void DateFormatPresets_GetStandardPresets_ReturnsValidFormattedDates()
        {
            var testDate = new DateTime(2026, 10, 28, 15, 30, 45);
            var presets = DateFormatPresets.GetStandardPresets(testDate);

            Assert.IsTrue(presets.Count >= 6);
            Assert.AreEqual("2026-10-28", DateFormatPresets.Format(testDate, "yyyy-MM-dd"));
            Assert.AreEqual("10/28/2026", DateFormatPresets.Format(testDate, "MM/dd/yyyy"));
            Assert.AreEqual("28/10/2026", DateFormatPresets.Format(testDate, "dd/MM/yyyy"));
            Assert.IsTrue(DateFormatPresets.Format(testDate, "MMMM dd, yyyy").Contains("October"));
        }

        [TestMethod]
        public void MockExcelContext_NamedRanges()
        {
            var ctx = new MockExcelContext();
            var names = ctx.GetNamedRanges();

            Assert.AreEqual(3, names.Count);
            Assert.IsTrue(names.Any(n => n.Name == "SalesTotal" && n.IsValid));
            Assert.IsTrue(names.Any(n => n.Name == "BrokenRef" && !n.IsValid));

            bool navigated = ctx.GoToNamedRange("SalesTotal");
            Assert.IsTrue(navigated);
            Assert.AreEqual("SalesTotal", ctx.LastNavigatedNamedRange);

            bool deleted = ctx.DeleteNamedRange("BrokenRef");
            Assert.IsTrue(deleted);
            Assert.AreEqual("BrokenRef", ctx.LastDeletedNamedRange);
        }

        [TestMethod]
        public void MockExcelContext_ColumnsManagement()
        {
            var ctx = new MockExcelContext();
            var cols = ctx.GetWorksheetColumns();

            Assert.AreEqual(4, cols.Count);
            Assert.AreEqual("A", cols[0].ColumnLetter);
            Assert.AreEqual("ID", cols[0].Header);
            Assert.IsFalse(cols[0].IsHidden);
            Assert.IsTrue(cols[2].IsHidden); // Department column

            ctx.SelectColumn(2);
            Assert.AreEqual(2, ctx.LastSelectedColumn);

            ctx.SetColumnsVisibility(new[] { 1, 2 }, true);
            Assert.AreEqual(1, ctx.ColumnVisibilityHistory.Count);
            Assert.IsTrue(ctx.ColumnVisibilityHistory[0].hidden);

            ctx.AutoFitColumns(new[] { 1, 2, 3 });
            Assert.AreEqual(1, ctx.AutoFittedColumnsHistory.Count);
            Assert.AreEqual(3, ctx.AutoFittedColumnsHistory[0].Count);
        }

        [TestMethod]
        public void MockExcelContext_CellDetailsAndContext()
        {
            var ctx = new MockExcelContext();
            var details = ctx.GetActiveCellDetails();

            Assert.IsNotNull(details);
            Assert.AreEqual("$B$5", details.Address);
            Assert.AreEqual("Sheet1", details.SheetName);
            Assert.AreEqual(2, details.LineCount); // Multi-line test string
            Assert.IsTrue(details.WordCount >= 5);
            Assert.AreEqual("Numeric", details.DataType);

            var cellCtx = ctx.GetActiveCellContext();
            Assert.AreEqual(1, cellCtx.row);
            Assert.AreEqual(1, cellCtx.col);

            ctx.HighlightActiveRowAndColumn(5, 3);
            Assert.AreEqual((5, 3), ctx.LastHighlightedRowAndCol);
        }

        [TestMethod]
        public void MockExcelContext_TableDataAndAppendRow()
        {
            var ctx = new MockExcelContext();
            var table = ctx.GetActiveTableData();

            Assert.IsNotNull(table);
            Assert.AreEqual(4, table.ColumnCount);
            Assert.AreEqual(2, table.RowCount);
            Assert.AreEqual("ID", table.Headers[0]);
            Assert.AreEqual("Salary", table.Headers[3]);

            var newRow = new object?[] { 103, "Charlie", "Marketing", 75000 };
            bool appended = ctx.AppendTableRow(newRow, table.StartRow, table.StartCol, table.ColumnCount);

            Assert.IsTrue(appended);
            Assert.AreEqual(1, ctx.AppendedRowsHistory.Count);
            Assert.AreEqual(103, ctx.AppendedRowsHistory[0].row[0]);
        }

        #region Batch 11: Formulas & Formatting Expansion Tests

        [TestMethod]
        public void FormulaBatchReplacer_ReplacesSheetReferences_PreservingStringLiterals()
        {
            var options = new FormulaReplaceOptions
            {
                FindText = "'OldSheet'!",
                ReplaceWith = "'2026_Final'!",
                MatchCase = false,
                MatchWholeWord = false,
                IgnoreStringLiterals = true
            };

            string formula = "=IF(A1=\"'OldSheet'!\", 'OldSheet'!B2 + 'OldSheet'!C2, 0)";
            string result = FormulaBatchReplacer.ReplaceInFormula(formula, options, out int count);

            Assert.AreEqual(2, count);
            Assert.IsTrue(result.Contains("\"'OldSheet'!\"")); // Quoted literal unchanged
            Assert.IsTrue(result.Contains("'2026_Final'!B2"));
            Assert.IsTrue(result.Contains("'2026_Final'!C2"));
        }

        [TestMethod]
        public void FormulaBatchReplacer_WholeWordOption_ProtectsPrefixWords()
        {
            var options = new FormulaReplaceOptions
            {
                FindText = "SUM",
                ReplaceWith = "AVERAGE",
                MatchCase = true,
                MatchWholeWord = true,
                IgnoreStringLiterals = true
            };

            string formula = "=SUM(A1:A10) + SUMPRODUCT(B1:B10, C1:C10)";
            string result = FormulaBatchReplacer.ReplaceInFormula(formula, options, out int count);

            Assert.AreEqual(1, count);
            Assert.IsTrue(result.StartsWith("=AVERAGE("));
            Assert.IsTrue(result.Contains("SUMPRODUCT(")); // Not changed
        }

        [TestMethod]
        public void FormulaBatchReplacer_ProcessFormulasMatrix_TracksChangesAccurately()
        {
            object?[,] matrix = new object?[2, 2]
            {
                { "=Sheet1!A1*2", "Static Text" },
                { 12345,          "=Sheet1!B10 + Sheet1!C10" }
            };

            var options = new FormulaReplaceOptions
            {
                FindText = "Sheet1!",
                ReplaceWith = "Budget!",
                MatchCase = false,
                MatchWholeWord = false
            };

            var res = FormulaBatchReplacer.ProcessFormulas(matrix, options, startRow: 1, startCol: 1);

            Assert.AreEqual(4, res.TotalCellsScanned);
            Assert.AreEqual(2, res.FormulasFound);
            Assert.AreEqual(2, res.FormulasModified);
            Assert.AreEqual(3, res.TotalReplacements);
            Assert.AreEqual(2, res.Changes.Count);

            Assert.AreEqual("=Budget!A1*2", res.UpdatedFormulas[0, 0]);
            Assert.AreEqual("Static Text", res.UpdatedFormulas[0, 1]);
            Assert.AreEqual(12345, res.UpdatedFormulas[1, 0]);
            Assert.AreEqual("=Budget!B10 + Budget!C10", res.UpdatedFormulas[1, 1]);
        }

        [TestMethod]
        public void FormulaTemplateLibrary_TemplatesLoadAndRenderCorrectly()
        {
            var templates = FormulaTemplateLibrary.GetAllTemplates();
            Assert.IsTrue(templates.Count >= 15);

            var cagr = templates.FirstOrDefault(t => t.Id == "fin_cagr");
            Assert.IsNotNull(cagr);
            Assert.AreEqual(FormulaCategory.Financial, cagr.Category);

            var customParams = new Dictionary<string, string>
            {
                { "EndValue", "D10" },
                { "StartValue", "D2" },
                { "Periods", "8" }
            };

            string rendered = cagr.RenderFormula(customParams);
            Assert.AreEqual("=((D10/D2)^(1/8))-1", rendered);

            var vlookup = templates.FirstOrDefault(t => t.Id == "lookup_vlookup_safe");
            Assert.IsNotNull(vlookup);
            string defRendered = vlookup.RenderFormula(null);
            Assert.IsTrue(defRendered.StartsWith("=IFERROR(VLOOKUP("));

            var searchResults = FormulaTemplateLibrary.Search("working days");
            Assert.IsTrue(searchResults.Count >= 1);
            Assert.AreEqual("date_workdays", searchResults[0].Id);
        }

        [TestMethod]
        public void NumberFormatPresets_BuildFormat_And_Preview_BehaveCorrectly()
        {
            var presets = NumberFormatPresets.GetAllPresets();
            Assert.IsTrue(presets.Count >= 10);

            // 1. Currency with Red Negative
            string currMask = NumberFormatPresets.BuildFormat(
                NumberFormatCategory.Currency,
                decimals: 2,
                useThousandSeparator: true,
                negativeStyle: NegativeNumberStyle.RedParentheses,
                currencySymbol: "$");

            Assert.AreEqual("$#,##0.00;[Red]($#,##0.00);$0.00", currMask);

            // 2. Percentage with 1 decimal
            string pctMask = NumberFormatPresets.BuildFormat(
                NumberFormatCategory.Percentage,
                decimals: 1,
                useThousandSeparator: false,
                negativeStyle: NegativeNumberStyle.StandardMinus);

            Assert.AreEqual("0.0%;-0.0%;0.0%", pctMask);

            // 3. Integer with commas and parentheses
            string intMask = NumberFormatPresets.BuildFormat(
                NumberFormatCategory.Integer,
                decimals: 0,
                useThousandSeparator: true,
                negativeStyle: NegativeNumberStyle.Parentheses);

            Assert.AreEqual("#,##0;(#,##0);0", intMask);

            // 4. Preview formatter
            string samplePos = NumberFormatPresets.FormatSamplePreview(1234.56, "$#,##0.00;($#,##0.00);$0.00");
            Assert.IsTrue(samplePos.Contains("1,234.56") || samplePos.Contains("1234.56"));
        }

        [TestMethod]
        public void DateFormatStandardizer_StandardizeRange_ParsesTextDates()
        {
            var standards = DateFormatStandardizer.GetStandards();
            Assert.IsTrue(standards.Count >= 8);

            object?[,] matrix = new object?[3, 2]
            {
                { "2026-10-28", "Not a date" },
                { "15/04/2026", 45000.0 }, // Serial date
                { "October 5, 2026", null }
            };

            var res = DateFormatStandardizer.StandardizeRange(matrix, "yyyy-mm-dd", convertTextDates: true);

            Assert.AreEqual(6, res.TotalCells);
            Assert.AreEqual(4, res.DateCellsFound);
            Assert.AreEqual(3, res.TextDatesConverted);
            Assert.AreEqual(2, res.UnchangedCells);

            // Verify the first cell is now a double serial date
            Assert.IsInstanceOfType(res.ProcessedValues[0, 0], typeof(double));
            Assert.AreEqual("Not a date", res.ProcessedValues[0, 1]);
        }

        [TestMethod]
        public void MockExcelContext_Batch11_PrecedentsDependentsAndFormatting()
        {
            var ctx = new MockExcelContext();

            // 1. Precedents
            var precs = ctx.GetActiveCellPrecedents(directOnly: true);
            Assert.AreEqual(3, precs.Count);
            Assert.AreEqual("Sheet1", precs[0].SheetName);
            Assert.AreEqual("$A$1", precs[0].CellAddress);
            Assert.AreEqual("'Sheet1'!$A$1", precs[0].FullAddress);

            // 2. Dependents
            var deps = ctx.GetActiveCellDependents(directOnly: true);
            Assert.AreEqual(2, deps.Count);
            Assert.AreEqual("$B$10", deps[0].CellAddress);
            Assert.IsTrue(deps[0].IsFormula);

            // 3. Navigation
            bool nav = ctx.SelectCellAddress("Sheet2", "$C$10");
            Assert.IsTrue(nav);
            Assert.AreEqual(("Sheet2", "$C$10"), ctx.LastNavigatedCell);

            // 4. Number Formatting
            ctx.ApplyNumberFormat("$#,##0.00", selectionOnly: true);
            Assert.AreEqual(1, ctx.AppliedNumberFormats.Count);
            Assert.AreEqual("$#,##0.00", ctx.AppliedNumberFormats[0].formatString);
            Assert.IsTrue(ctx.AppliedNumberFormats[0].selectionOnly);

            // 5. Formula Insertion
            ctx.InsertFormulaAtActiveCell("=SUM(A1:A10)");
            Assert.AreEqual(1, ctx.InsertedFormulasHistory.Count);
            Assert.AreEqual("=SUM(A1:A10)", ctx.InsertedFormulasHistory[0]);
        }

        [TestMethod]
        public void Batch11_Commands_CanBeRegisteredAndExecuted()
        {
            var registry = new CommandRegistry();
            var ctx = new MockExcelContext();

            var cmd1 = new SuiteCommand(
                id: "CMD_FORMULA_AUDIT_PRECEDENTS",
                title: "Audit Precedents",
                description: "desc",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: c => c.ShowMessage("Precedents", "Audit"),
                keywords: new[] { "precedents" },
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 11");

            var cmd2 = new SuiteCommand(
                id: "CMD_FORMULA_AUDIT_DEPENDENTS",
                title: "Audit Dependents",
                description: "desc",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: c => c.ShowMessage("Dependents", "Audit"),
                keywords: new[] { "dependents" },
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 11");

            var cmd3 = new SuiteCommand(
                id: "CMD_FORMULA_BATCH_REPLACE_REF",
                title: "Replace Refs",
                description: "desc",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: c => c.ShowMessage("Replace", "Refs"),
                keywords: new[] { "replace" },
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 11");

            var cmd4 = new SuiteCommand(
                id: "CMD_FORMULA_TEMPLATES",
                title: "Templates",
                description: "desc",
                category: CommandCategory.FormulasAndCalculation,
                executeAction: c => c.ShowMessage("Templates", "Lib"),
                keywords: new[] { "templates" },
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "G",
                roadmapBatch: "Batch 11");

            var cmd5 = new SuiteCommand(
                id: "CMD_FORMAT_BATCH_NUMBERS",
                title: "Format Numbers",
                description: "desc",
                category: CommandCategory.FormattingAndDesign,
                executeAction: c => c.ApplyNumberFormat("$#,##0.00"),
                keywords: new[] { "numbers" },
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "H",
                roadmapBatch: "Batch 11");

            var cmd6 = new SuiteCommand(
                id: "CMD_FORMAT_DATE_STANDARDS",
                title: "Date Standards",
                description: "desc",
                category: CommandCategory.FormattingAndDesign,
                executeAction: c => c.ApplyNumberFormat("yyyy-mm-dd"),
                keywords: new[] { "dates" },
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "H",
                roadmapBatch: "Batch 11");

            registry.RegisterCommand(cmd1);
            registry.RegisterCommand(cmd2);
            registry.RegisterCommand(cmd3);
            registry.RegisterCommand(cmd4);
            registry.RegisterCommand(cmd5);
            registry.RegisterCommand(cmd6);

            Assert.AreEqual(6, registry.GetAllCommands().Count);
            registry.Execute("CMD_FORMAT_BATCH_NUMBERS", ctx);
            Assert.AreEqual(1, ctx.AppliedNumberFormats.Count);
            Assert.AreEqual("$#,##0.00", ctx.AppliedNumberFormats[0].formatString);

            registry.Execute("CMD_FORMULA_AUDIT_PRECEDENTS", ctx);
            Assert.IsTrue(ctx.MessagesShown.Any(m => m.Contains("Precedents")));
        }

        #endregion

        #endregion

        #endregion

        #endregion
    }
}
