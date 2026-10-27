using System;
using ExcelProductivitySuite.Engine.DataCleaning;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class RecordComparerTests
    {
        [TestMethod]
        public void CompareRecords_IdentifiesSourceOnly_TargetOnly_AndMatching()
        {
            // Source: Header + 3 rows (ID 101, 102, 103)
            object?[,] source = new object?[4, 2]
            {
                { "ID", "Name" },
                { 101, "Alice" },
                { 102, "Bob" },
                { 103, "Charlie" }
            };

            // Target: Header + 3 rows (ID 102, 103, 104)
            object?[,] target = new object?[4, 2]
            {
                { "ID", "Department" },
                { 102, "HR" },
                { 103, "Engineering" },
                { 104, "Sales" }
            };

            var options = new RecordCompareOptions
            {
                HasHeaders = true,
                SourceKeyColumns = new[] { 0 },
                TargetKeyColumns = new[] { 0 },
                MatchCase = false
            };

            var result = RangeRecordComparer.CompareRecords(source, target, options);

            // SourceOnly: ID 101
            Assert.AreEqual(1, result.SourceOnly.Count);
            Assert.AreEqual("101", result.SourceOnly[0].CompositeKey);

            // TargetOnly: ID 104
            Assert.AreEqual(1, result.TargetOnly.Count);
            Assert.AreEqual("104", result.TargetOnly[0].CompositeKey);

            // Matching: ID 102, 103
            Assert.AreEqual(2, result.Matching.Count);

            // Verify matrix export
            var exported = result.ExportToMatrix(RecordResultCategory.SourceOnly);
            Assert.AreEqual(2, exported.GetLength(0)); // 1 header + 1 row
            Assert.AreEqual("ID", exported[0, 0]);
            Assert.AreEqual(101, exported[1, 0]);
        }

        [TestMethod]
        public void CompareRecords_CompositeKey_MultipleColumnsSupported()
        {
            object?[,] source = new object?[3, 3]
            {
                { "First", "Last", "City" },
                { "John", "Doe", "New York" },
                { "Jane", "Smith", "Chicago" }
            };

            object?[,] target = new object?[3, 3]
            {
                { "FirstName", "LastName", "Role" },
                { "John", "Doe", "Manager" },
                { "Bob", "Smith", "Analyst" }
            };

            var options = new RecordCompareOptions
            {
                HasHeaders = true,
                SourceKeyColumns = new[] { 0, 1 }, // First + Last
                TargetKeyColumns = new[] { 0, 1 },
                MatchCase = false
            };

            var result = RangeRecordComparer.CompareRecords(source, target, options);

            Assert.AreEqual(1, result.Matching.Count);
            Assert.AreEqual("John│Doe", result.Matching[0].CompositeKey);
            Assert.AreEqual(1, result.SourceOnly.Count);
            Assert.AreEqual("Jane│Smith", result.SourceOnly[0].CompositeKey);
            Assert.AreEqual(1, result.TargetOnly.Count);
            Assert.AreEqual("Bob│Smith", result.TargetOnly[0].CompositeKey);
        }

        [TestMethod]
        public void CompareRecords_InternalDuplicates_DetectedAndCategorized()
        {
            object?[,] source = new object?[4, 2]
            {
                { "ID", "Name" },
                { 101, "Alice" },
                { 101, "Alice Duplicate" },
                { 102, "Bob" }
            };

            object?[,] target = new object?[2, 2]
            {
                { "ID", "Dept" },
                { 102, "Finance" }
            };

            var options = new RecordCompareOptions
            {
                HasHeaders = true,
                SourceKeyColumns = new[] { 0 },
                TargetKeyColumns = new[] { 0 }
            };

            var result = RangeRecordComparer.CompareRecords(source, target, options);

            Assert.AreEqual(1, result.SourceDuplicates.Count);
            Assert.AreEqual("101", result.SourceDuplicates[0].CompositeKey);
        }
    }
}
