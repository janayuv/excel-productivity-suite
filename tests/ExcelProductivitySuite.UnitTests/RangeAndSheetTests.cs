using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.Engine.Ranges;
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
    }
}
