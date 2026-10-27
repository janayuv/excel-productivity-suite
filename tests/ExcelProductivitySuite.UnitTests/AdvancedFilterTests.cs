using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Engine.DataCleaning;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class AdvancedFilterTests
    {
        [TestMethod]
        public void Filter_AndLogic_FiltersCorrectly()
        {
            object?[,] matrix = new object?[5, 3]
            {
                { "Name", "Department", "Salary" },
                { "Alice", "Engineering", 90000 },
                { "Bob", "Sales", 60000 },
                { "Charlie", "Engineering", 50000 },
                { "Dana", "Engineering", 120000 }
            };

            var options = new AdvancedFilterOptions
            {
                Conjunction = FilterConjunction.And,
                HasHeaders = true,
                Conditions = new List<FilterCondition>
                {
                    new FilterCondition { ColumnIndex = 1, Operator = FilterOperator.Equals, ComparisonValue = "Engineering" },
                    new FilterCondition { ColumnIndex = 2, Operator = FilterOperator.GreaterThan, ComparisonValue = "70000" }
                }
            };

            var result = AdvancedFilter.Execute(matrix, options);

            // Alice (90000) and Dana (120000) match
            Assert.AreEqual(2, result.MatchCount);
            Assert.AreEqual("Alice", result.MatchingRows[0][0]);
            Assert.AreEqual("Dana", result.MatchingRows[1][0]);

            var exported = result.ExportToMatrix();
            Assert.AreEqual(3, exported.GetLength(0)); // Header + 2 rows
            Assert.AreEqual("Name", exported[0, 0]);
        }

        [TestMethod]
        public void Filter_OrLogic_MatchesAnyCondition()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { "Product", "Status" },
                { "Widget A", "Pending" },
                { "Widget B", "Shipped" },
                { "Widget C", "Cancelled" }
            };

            var options = new AdvancedFilterOptions
            {
                Conjunction = FilterConjunction.Or,
                HasHeaders = true,
                Conditions = new List<FilterCondition>
                {
                    new FilterCondition { ColumnIndex = 1, Operator = FilterOperator.Equals, ComparisonValue = "Pending" },
                    new FilterCondition { ColumnIndex = 1, Operator = FilterOperator.Equals, ComparisonValue = "Cancelled" }
                }
            };

            var result = AdvancedFilter.Execute(matrix, options);

            Assert.AreEqual(2, result.MatchCount);
            Assert.AreEqual("Widget A", result.MatchingRows[0][0]);
            Assert.AreEqual("Widget C", result.MatchingRows[1][0]);
        }
    }
}
