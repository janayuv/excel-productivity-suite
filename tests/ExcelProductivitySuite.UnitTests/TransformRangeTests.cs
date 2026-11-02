using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Ranges;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class TransformRangeTests
    {
        [TestMethod]
        public void SingleColumnToRange_FixedColumns_GroupsProperly()
        {
            // Arrange: 10 items in a single column
            var source = new object?[10, 1]
            {
                { "A" }, { "B" }, { "C" },
                { "D" }, { "E" }, { "F" },
                { "G" }, { "H" }, { "I" },
                { "J" }
            };

            var options = new TransformRangeOptions
            {
                Type = TransformType.SingleColumnToRange,
                RecordSplit = TransformRecordSplit.FixedValue,
                ColumnsPerRecord = 3
            };

            // Act
            var result = RangeTransformer.Transform(source, options);

            // Assert: ceil(10 / 3) = 4 rows, 3 columns
            Assert.AreEqual(4, result.GetLength(0));
            Assert.AreEqual(3, result.GetLength(1));

            Assert.AreEqual("A", result[0, 0]);
            Assert.AreEqual("B", result[0, 1]);
            Assert.AreEqual("C", result[0, 2]);

            Assert.AreEqual("D", result[1, 0]);
            Assert.AreEqual("E", result[1, 1]);
            Assert.AreEqual("F", result[1, 2]);

            Assert.AreEqual("G", result[2, 0]);
            Assert.AreEqual("H", result[2, 1]);
            Assert.AreEqual("I", result[2, 2]);

            Assert.AreEqual("J", result[3, 0]);
            Assert.IsNull(result[3, 1]);
            Assert.IsNull(result[3, 2]);
        }

        [TestMethod]
        public void SingleColumnToRange_BlankCellDelimited_SplitsOnBlanks()
        {
            // Arrange: 2 contact records separated by blank cell
            var source = new object?[8, 1]
            {
                { "Alice" },
                { "555-0100" },
                { "alice@example.com" },
                { null }, // separator
                { "Bob" },
                { "555-0200" },
                { "bob@example.com" },
                { "Manager" } // 4 items in second record
            };

            var options = new TransformRangeOptions
            {
                Type = TransformType.SingleColumnToRange,
                RecordSplit = TransformRecordSplit.BlankCellDelimited
            };

            // Act
            var result = RangeTransformer.Transform(source, options);

            // Assert: 2 rows, max cols = 4
            Assert.AreEqual(2, result.GetLength(0));
            Assert.AreEqual(4, result.GetLength(1));

            Assert.AreEqual("Alice", result[0, 0]);
            Assert.AreEqual("555-0100", result[0, 1]);
            Assert.AreEqual("alice@example.com", result[0, 2]);
            Assert.IsNull(result[0, 3]);

            Assert.AreEqual("Bob", result[1, 0]);
            Assert.AreEqual("555-0200", result[1, 1]);
            Assert.AreEqual("bob@example.com", result[1, 2]);
            Assert.AreEqual("Manager", result[1, 3]);
        }

        [TestMethod]
        public void SingleRowToRange_FixedColumns_GroupsProperly()
        {
            // Arrange: 8 items in a single row
            var source = new object?[1, 8]
            {
                { 1, 2, 3, 4, 5, 6, 7, 8 }
            };

            var options = new TransformRangeOptions
            {
                Type = TransformType.SingleRowToRange,
                RecordSplit = TransformRecordSplit.FixedValue,
                ColumnsPerRecord = 4
            };

            // Act
            var result = RangeTransformer.Transform(source, options);

            // Assert: 2 rows, 4 columns
            Assert.AreEqual(2, result.GetLength(0));
            Assert.AreEqual(4, result.GetLength(1));

            Assert.AreEqual(1, result[0, 0]);
            Assert.AreEqual(2, result[0, 1]);
            Assert.AreEqual(3, result[0, 2]);
            Assert.AreEqual(4, result[0, 3]);

            Assert.AreEqual(5, result[1, 0]);
            Assert.AreEqual(6, result[1, 1]);
            Assert.AreEqual(7, result[1, 2]);
            Assert.AreEqual(8, result[1, 3]);
        }

        [TestMethod]
        public void RangeToSingleColumn_ByRow_FlattensCorrectly()
        {
            var source = new object?[2, 3]
            {
                { "R1C1", "R1C2", "R1C3" },
                { "R2C1", "R2C2", "R2C3" }
            };

            var options = new TransformRangeOptions
            {
                Type = TransformType.RangeToSingleColumn,
                Order = TransformOrder.ByRow,
                SkipBlankCells = false
            };

            var result = RangeTransformer.Transform(source, options);

            Assert.AreEqual(6, result.GetLength(0));
            Assert.AreEqual(1, result.GetLength(1));

            Assert.AreEqual("R1C1", result[0, 0]);
            Assert.AreEqual("R1C2", result[1, 0]);
            Assert.AreEqual("R1C3", result[2, 0]);
            Assert.AreEqual("R2C1", result[3, 0]);
            Assert.AreEqual("R2C2", result[4, 0]);
            Assert.AreEqual("R2C3", result[5, 0]);
        }

        [TestMethod]
        public void RangeToSingleColumn_ByColumn_FlattensCorrectly()
        {
            var source = new object?[2, 3]
            {
                { "R1C1", "R1C2", "R1C3" },
                { "R2C1", "R2C2", "R2C3" }
            };

            var options = new TransformRangeOptions
            {
                Type = TransformType.RangeToSingleColumn,
                Order = TransformOrder.ByColumn,
                SkipBlankCells = false
            };

            var result = RangeTransformer.Transform(source, options);

            Assert.AreEqual(6, result.GetLength(0));
            Assert.AreEqual(1, result.GetLength(1));

            Assert.AreEqual("R1C1", result[0, 0]);
            Assert.AreEqual("R2C1", result[1, 0]);
            Assert.AreEqual("R1C2", result[2, 0]);
            Assert.AreEqual("R2C2", result[3, 0]);
            Assert.AreEqual("R1C3", result[4, 0]);
            Assert.AreEqual("R2C3", result[5, 0]);
        }

        [TestMethod]
        public void RangeToSingleColumn_SkipBlanks_FiltersOutBlanksAndEmptyStrings()
        {
            var source = new object?[3, 2]
            {
                { "Alpha", null },
                { "   ", "Beta" },
                { "", "Gamma" }
            };

            var options = new TransformRangeOptions
            {
                Type = TransformType.RangeToSingleColumn,
                Order = TransformOrder.ByRow,
                SkipBlankCells = true
            };

            var result = RangeTransformer.Transform(source, options);

            Assert.AreEqual(3, result.GetLength(0));
            Assert.AreEqual(1, result.GetLength(1));

            Assert.AreEqual("Alpha", result[0, 0]);
            Assert.AreEqual("Beta", result[1, 0]);
            Assert.AreEqual("Gamma", result[2, 0]);
        }

        [TestMethod]
        public void RangeToSingleRow_FlattensCorrectly()
        {
            var source = new object?[2, 2]
            {
                { "X1", "X2" },
                { "Y1", "Y2" }
            };

            var options = new TransformRangeOptions
            {
                Type = TransformType.RangeToSingleRow,
                Order = TransformOrder.ByRow,
                SkipBlankCells = false
            };

            var result = RangeTransformer.Transform(source, options);

            Assert.AreEqual(1, result.GetLength(0));
            Assert.AreEqual(4, result.GetLength(1));

            Assert.AreEqual("X1", result[0, 0]);
            Assert.AreEqual("X2", result[0, 1]);
            Assert.AreEqual("Y1", result[0, 2]);
            Assert.AreEqual("Y2", result[0, 3]);
        }

        [TestMethod]
        public void Transform_Handles1BasedComMatrixCorrectly()
        {
            // Simulate 1-based 2D array returned by Excel COM Interop
            Array comArray = Array.CreateInstance(typeof(object), new int[] { 3, 2 }, new int[] { 1, 1 });
            comArray.SetValue("C1R1", 1, 1);
            comArray.SetValue("C2R1", 1, 2);
            comArray.SetValue("C1R2", 2, 1);
            comArray.SetValue("C2R2", 2, 2);
            comArray.SetValue("C1R3", 3, 1);
            comArray.SetValue("C2R3", 3, 2);

            var options = new TransformRangeOptions
            {
                Type = TransformType.RangeToSingleColumn,
                Order = TransformOrder.ByRow,
                SkipBlankCells = false
            };

            var result = RangeTransformer.Transform((object?[,])comArray, options);

            Assert.AreEqual(6, result.GetLength(0));
            Assert.AreEqual(1, result.GetLength(1));
            Assert.AreEqual("C1R1", result[0, 0]);
            Assert.AreEqual("C2R1", result[1, 0]);
            Assert.AreEqual("C1R2", result[2, 0]);
            Assert.AreEqual("C2R2", result[3, 0]);
            Assert.AreEqual("C1R3", result[4, 0]);
            Assert.AreEqual("C2R3", result[5, 0]);
        }

        [TestMethod]
        public void OptionsValidation_ChecksInvalidInputs()
        {
            var options = new TransformRangeOptions
            {
                Type = TransformType.SingleColumnToRange,
                RecordSplit = TransformRecordSplit.FixedValue,
                ColumnsPerRecord = 0
            };

            var (isValid, err) = options.Validate();
            Assert.IsFalse(isValid);
            StringAssert.Contains(err, "greater than or equal to 1");

            options.ColumnsPerRecord = 3;
            options.DestinationMode = TransformDestinationMode.TargetCell;
            options.TargetAddress = "";
            (isValid, err) = options.Validate();
            Assert.IsFalse(isValid);
            StringAssert.Contains(err, "destination cell address");

            options.DestinationMode = TransformDestinationMode.NewSheet;
            options.NewSheetName = "   ";
            (isValid, err) = options.Validate();
            Assert.IsFalse(isValid);
            StringAssert.Contains(err, "new worksheet");
        }

        [TestMethod]
        public void MockExcelContext_TransformRange_InPlace_And_NewSheet()
        {
            var mock = new MockExcelContext();
            mock.CurrentValues = new object?[6, 1]
            {
                { "Item 1" },
                { "Item 2" },
                { "Item 3" },
                { "Item 4" },
                { "Item 5" },
                { "Item 6" }
            };

            var options = new TransformRangeOptions
            {
                Type = TransformType.SingleColumnToRange,
                RecordSplit = TransformRecordSplit.FixedValue,
                ColumnsPerRecord = 2,
                DestinationMode = TransformDestinationMode.InPlace
            };

            var res = mock.TransformRange(options);
            Assert.IsTrue(res.Success);
            Assert.IsNotNull(mock.CurrentValues);
            Assert.AreEqual(3, mock.CurrentValues!.GetLength(0));
            Assert.AreEqual(2, mock.CurrentValues.GetLength(1));

            // Test New Sheet
            options.DestinationMode = TransformDestinationMode.NewSheet;
            options.NewSheetName = "TransformedReport";
            var res2 = mock.TransformRange(options);
            Assert.IsTrue(res2.Success);
            Assert.IsTrue(mock.CreatedSheets.ContainsKey("TransformedReport"));
        }
    }
}
