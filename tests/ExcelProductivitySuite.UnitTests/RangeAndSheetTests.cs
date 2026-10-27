using System;
using System.Collections.Generic;
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
    }
}
