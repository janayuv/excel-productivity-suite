using System;
using ExcelProductivitySuite.Engine.Ranges;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class RangeMergerTests
    {
        [TestMethod]
        public void Merge_AppendWithHeaders_AlignsColumnsByHeaderName()
        {
            // Range 1: Name, Age
            object?[,] r1 = new object?[3, 2]
            {
                { "Name", "Age" },
                { "Alice", 30 },
                { "Bob", 25 }
            };

            // Range 2: Age, Name, City (Different column order + extra column)
            object?[,] r2 = new object?[2, 3]
            {
                { "Age", "Name", "City" },
                { 40, "Charlie", "Seattle" }
            };

            var options = new RangeMergeOptions
            {
                Mode = MergeMode.AppendRows,
                HasHeaders = true,
                AlignByHeaderNames = true,
                IncludeHeadersOnce = true
            };

            var result = RangeMerger.Merge(r1, r2, options);

            // Columns: Name, Age, City (3 columns total)
            Assert.AreEqual(3, result.TotalColumns);
            // Rows: Header + 2 from r1 + 1 from r2 = 4 rows
            Assert.AreEqual(4, result.TotalRows);

            // Check header
            Assert.AreEqual("Name", result.MergedMatrix[0, 0]);
            Assert.AreEqual("Age", result.MergedMatrix[0, 1]);
            Assert.AreEqual("City", result.MergedMatrix[0, 2]);

            // Charlie row aligned correctly
            Assert.AreEqual("Charlie", result.MergedMatrix[3, 0]);
            Assert.AreEqual(40, result.MergedMatrix[3, 1]);
            Assert.AreEqual("Seattle", result.MergedMatrix[3, 2]);

            // Alice City should be null/empty
            Assert.IsNull(result.MergedMatrix[1, 2]);
        }

        [TestMethod]
        public void Merge_SideBySide_JoinsHorizontally()
        {
            object?[,] r1 = new object?[2, 2]
            {
                { "A1", "B1" },
                { "A2", "B2" }
            };

            object?[,] r2 = new object?[2, 1]
            {
                { "C1" },
                { "C2" }
            };

            var options = new RangeMergeOptions { Mode = MergeMode.SideBySide };
            var result = RangeMerger.Merge(r1, r2, options);

            Assert.AreEqual(2, result.TotalRows);
            Assert.AreEqual(3, result.TotalColumns);
            Assert.AreEqual("A1", result.MergedMatrix[0, 0]);
            Assert.AreEqual("B1", result.MergedMatrix[0, 1]);
            Assert.AreEqual("C1", result.MergedMatrix[0, 2]);
        }
    }
}
