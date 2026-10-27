using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Engine.Worksheets;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class DataSplitterTests
    {
        [TestMethod]
        public void Split_ByDepartmentColumn_CreatesSeparatedSheetsWithHeaders()
        {
            // Table with Header + 4 data rows
            object?[,] matrix = new object?[5, 3]
            {
                { "ID", "Name", "Department" },
                { 101, "Alice", "Sales" },
                { 102, "Bob", "Engineering" },
                { 103, "Charlie", "Sales" },
                { 104, "Dana", "Engineering" }
            };

            var options = new DataSplitOptions
            {
                SplitColumnIndex = 2, // Department column
                HasHeaders = true
            };

            var result = DataSplitter.Split(matrix, options);

            Assert.AreEqual(4, result.TotalRowsProcessed);
            Assert.AreEqual(2, result.DistinctGroupsCount);
            Assert.IsTrue(result.SheetsToCreate.ContainsKey("Sales"));
            Assert.IsTrue(result.SheetsToCreate.ContainsKey("Engineering"));

            // Sales sheet has Header + 2 rows = 3 rows
            var salesMatrix = result.SheetsToCreate["Sales"];
            Assert.AreEqual(3, salesMatrix.GetLength(0));
            Assert.AreEqual(3, salesMatrix.GetLength(1));
            Assert.AreEqual("ID", salesMatrix[0, 0]);
            Assert.AreEqual(101, salesMatrix[1, 0]);
            Assert.AreEqual(103, salesMatrix[2, 0]);
        }

        [TestMethod]
        public void SanitizeSheetName_ForbiddenCharsAndLength_CleanedSafely()
        {
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Report_Q1" };

            // Forbidden chars
            string s1 = DataSplitter.SanitizeSheetName("US/West:Q1 [Final]", null);
            Assert.AreEqual("US_West_Q1 _Final_", s1);

            // Reserved name History
            string s2 = DataSplitter.SanitizeSheetName("History", null);
            Assert.AreEqual("History_Data", s2);

            // Length > 31
            string longStr = new string('A', 40);
            string s3 = DataSplitter.SanitizeSheetName(longStr, null);
            Assert.AreEqual(31, s3.Length);

            // Name collision uniquification
            string s4 = DataSplitter.SanitizeSheetName("Report_Q1", existing);
            Assert.AreEqual("Report_Q1_1", s4);
        }
    }
}
