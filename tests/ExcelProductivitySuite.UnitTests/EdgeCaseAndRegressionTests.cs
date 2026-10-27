using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.Engine.Worksheets;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class EdgeCaseAndRegressionTests
    {
        [TestMethod]
        public void Extract_DelimitersAtEdges_HandlesGracefully()
        {
            var options = new TextExtractionOptions
            {
                Mode = ExtractionMode.BeforeDelimiter,
                PrimaryDelimiter = ","
            };

            // Delimiter at start
            string res1 = TextExtractor.Extract(",LeadingComma", options);
            Assert.AreEqual("", res1);

            // Delimiter at end
            options.Mode = ExtractionMode.AfterDelimiter;
            string res2 = TextExtractor.Extract("TrailingComma,", options);
            Assert.AreEqual("", res2);
        }

        [TestMethod]
        public void Split_ConsecutiveDelimiters_HandlesEmptyTokens()
        {
            var options = new TextSplitOptions
            {
                Delimiter = ",",
                TrimTokens = true,
                RemoveEmptyTokens = false
            };

            string[] tokens = TextSplitter.Split("A,,,B", options);
            Assert.AreEqual(4, tokens.Length);
            Assert.AreEqual("A", tokens[0]);
            Assert.AreEqual("", tokens[1]);
            Assert.AreEqual("", tokens[2]);
            Assert.AreEqual("B", tokens[3]);
        }

        [TestMethod]
        public void Unicode_ChineseArabicAndEmoji_PreservedCorrectly()
        {
            // Unicode characters removal
            var charOptions = new CharacterRemovalOptions { RemoveNumeric = true };
            string resChinese = CharacterRemover.RemoveCharacters("订单编号 123456 完成", charOptions);
            Assert.AreEqual("订单编号  完成", resChinese);

            string resArabic = CharacterRemover.RemoveCharacters("طلب رقم 789 مؤكد", charOptions);
            Assert.AreEqual("طلب رقم  مؤكد", resArabic);

            string resEmoji = CharacterRemover.RemoveCharacters("Package 📦 99 delivered", charOptions);
            Assert.AreEqual("Package 📦  delivered", resEmoji);
        }

        [TestMethod]
        public void Whitespace_NonBreakingSpacesAndTabs_CleanedProperly()
        {
            // \u00A0 is non-breaking space (NBSP)
            string input = "\u00A0\u00A0Hello\t\tWorld\u00A0 \r\n";
            var options = new WhitespaceOptions
            {
                Mode = WhitespaceTrimMode.CollapseExcessInternalSpaces,
                ConvertNonBreakingSpaces = true
            };

            string result = WhitespaceNormalizer.Normalize(input, options);
            Assert.AreEqual("Hello World", result);
        }

        [TestMethod]
        public void PositionInserter_StartAndEndOffsets_InsertsCorrectly()
        {
            // Insert at position 1 from start
            var opts1 = new PositionInsertOptions
            {
                TextToInsert = "PRE-",
                PositionIndex = 1,
                FromEnd = false
            };
            string res1 = PositionInserter.Insert("12345", opts1);
            Assert.AreEqual("PRE-12345", res1);

            // Insert at position 3
            var opts2 = new PositionInsertOptions
            {
                TextToInsert = "-",
                PositionIndex = 3,
                FromEnd = false
            };
            string res2 = PositionInserter.Insert("12345", opts2);
            Assert.AreEqual("12-345", res2);

            // Insert from end
            var opts3 = new PositionInsertOptions
            {
                TextToInsert = "-SUF",
                PositionIndex = 0,
                FromEnd = true
            };
            string res3 = PositionInserter.Insert("12345", opts3);
            Assert.AreEqual("12345-SUF", res3);

            // Large offset beyond string length
            var opts4 = new PositionInsertOptions
            {
                TextToInsert = "END",
                PositionIndex = 999,
                FromEnd = false
            };
            string res4 = PositionInserter.Insert("123", opts4);
            Assert.AreEqual("123END", res4);
        }

        [TestMethod]
        public void FillBlanks_MixedDataTypes_PreservedCorrectly()
        {
            DateTime now = DateTime.Today;
            object?[,] matrix = new object?[4, 2];
            matrix[0, 0] = 100;
            matrix[1, 0] = null;
            matrix[2, 0] = null;
            matrix[3, 0] = 200;

            matrix[0, 1] = now;
            matrix[1, 1] = null;
            matrix[2, 1] = "Text";
            matrix[3, 1] = null;

            var options = new FillBlankOptions { Direction = FillBlankDirection.Down };
            var filled = RangeFill.FillBlanks(matrix, options, out int count);

            Assert.AreEqual(4, count);
            Assert.AreEqual(100, filled[1, 0]);
            Assert.AreEqual(100, filled[2, 0]);
            Assert.AreEqual(now, filled[1, 1]);
            Assert.AreEqual("Text", filled[3, 1]);
        }

        [TestMethod]
        public void DuplicateFinder_WithHeaders_DoesNotFlagHeaderAsDuplicate()
        {
            // Row 0 is header. Rows 1 and 2 have identical values to header ("Name", "Role")
            object?[,] matrix = new object?[3, 2];
            matrix[0, 0] = "Name"; matrix[0, 1] = "Role";
            matrix[1, 0] = "Name"; matrix[1, 1] = "Role";
            matrix[2, 0] = "Name"; matrix[2, 1] = "Role";

            var options = new DuplicateDetectionOptions
            {
                HasHeaders = true,
                KeepStrategy = DuplicateKeepStrategy.KeepFirst
            };

            bool[] isDuplicate = DuplicateFinder.IdentifyDuplicateRows(matrix, options);

            // Row 0 is header -> must NOT be duplicate
            Assert.IsFalse(isDuplicate[0]);
            // Row 1 is first data row -> must NOT be duplicate
            Assert.IsFalse(isDuplicate[1]);
            // Row 2 is duplicate of row 1 -> MUST be duplicate
            Assert.IsTrue(isDuplicate[2]);
        }

        [TestMethod]
        public void DuplicateFinder_SpecificKeyColumns_EvaluatesOnlySelectedColumns()
        {
            object?[,] matrix = new object?[3, 3];
            // Row 0: ID=1, Name="Alice", City="NYC"
            matrix[0, 0] = 1; matrix[0, 1] = "Alice"; matrix[0, 2] = "NYC";
            // Row 1: ID=1, Name="Bob", City="LA" (same ID, different name/city)
            matrix[1, 0] = 1; matrix[1, 1] = "Bob"; matrix[1, 2] = "LA";
            // Row 2: ID=2, Name="Alice", City="NYC"
            matrix[2, 0] = 2; matrix[2, 1] = "Alice"; matrix[2, 2] = "NYC";

            var options = new DuplicateDetectionOptions
            {
                HasHeaders = false,
                KeepStrategy = DuplicateKeepStrategy.KeepFirst,
                KeyColumnIndices = new[] { 0 } // Compare only column 0 (ID)
            };

            bool[] isDuplicate = DuplicateFinder.IdentifyDuplicateRows(matrix, options);

            Assert.IsFalse(isDuplicate[0]);
            Assert.IsTrue(isDuplicate[1]);  // Duplicate ID 1
            Assert.IsFalse(isDuplicate[2]); // Unique ID 2
        }

        [TestMethod]
        public void WorksheetNameValidator_ReservedHistoryName_Rejected()
        {
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Sheet1" };
            var result = WorksheetNameValidator.ValidateSheetName("History", existing);

            Assert.IsFalse(result.IsValid);
            Assert.IsTrue(result.ErrorMessage!.Contains("reserved"));
        }

        [TestMethod]
        public void WorksheetNameValidator_MaxLength31Characters_Enforced()
        {
            string longName = new string('A', 32);
            var result = WorksheetNameValidator.ValidateSheetName(longName, new HashSet<string>());

            Assert.IsFalse(result.IsValid);
            Assert.IsTrue(result.ErrorMessage!.Contains("31"));
        }

        [TestMethod]
        public void RangeComparer_CellByCell_DetectsMismatchesAccurately()
        {
            object?[,] source = new object?[2, 2]
            {
                { "A", 100 },
                { "B", 200 }
            };

            object?[,] target = new object?[2, 2]
            {
                { "A", 100 },
                { "B", 999 } // Mismatch at (2, 2)
            };

            var options = new RangeCompareOptions
            {
                Mode = RangeCompareMode.CellByCell,
                MatchCase = true
            };

            var result = RangeComparer.Compare(source, target, options);

            Assert.AreEqual(4, result.TotalCellsCompared);
            Assert.AreEqual(1, result.TotalDifferences);
            Assert.AreEqual(2, result.Differences[0].Row);
            Assert.AreEqual(2, result.Differences[0].Column);
            Assert.AreEqual("200", result.Differences[0].SourceValue);
            Assert.AreEqual("999", result.Differences[0].TargetValue);
        }

        [TestMethod]
        public void RangeComparer_FindMissingValues_IdentifiesMissingItems()
        {
            object?[,] source = new object?[3, 1] { { "Apple" }, { "Banana" }, { "Orange" } };
            object?[,] target = new object?[2, 1] { { "Apple" }, { "Grape" } };

            var options = new RangeCompareOptions
            {
                Mode = RangeCompareMode.FindValuesNotInTarget,
                MatchCase = false
            };

            var result = RangeComparer.Compare(source, target, options);

            Assert.AreEqual(2, result.MissingValues.Count);
            Assert.IsTrue(result.MissingValues.Contains("Banana"));
            Assert.IsTrue(result.MissingValues.Contains("Orange"));
        }

        [TestMethod]
        public void RangeSelector_FindErrorsAndMaxMin_WorksCorrectly()
        {
            object?[,] matrix = new object?[4, 1]
            {
                { 15.5 },
                { "#DIV/0!" },
                { 99.9 },
                { 3.2 }
            };

            // Find errors
            var errorCells = RangeSelector.FindMatchingCells(matrix, 1, 1, SelectionCriteria.ErrorCells);
            Assert.AreEqual(1, errorCells.Count);
            Assert.AreEqual(2, errorCells[0].Row);
            Assert.AreEqual("#DIV/0!", errorCells[0].Value);

            // Find max
            var maxCells = RangeSelector.FindMatchingCells(matrix, 1, 1, SelectionCriteria.MaxNumericValue);
            Assert.AreEqual(1, maxCells.Count);
            Assert.AreEqual(3, maxCells[0].Row);
            Assert.AreEqual(99.9, (double)maxCells[0].Value!);

            // Find min
            var minCells = RangeSelector.FindMatchingCells(matrix, 1, 1, SelectionCriteria.MinNumericValue);
            Assert.AreEqual(1, minCells.Count);
            Assert.AreEqual(4, minCells[0].Row);
            Assert.AreEqual(3.2, (double)minCells[0].Value!);
        }
    }
}
