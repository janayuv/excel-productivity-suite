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
    public class EngineReliabilityAndLeakTests
    {
        [TestMethod]
        public void RepeatedExecution_TextExtractor_ZeroStateLeaksAcross1000Iterations()
        {
            var opts = new TextExtractionOptions
            {
                Mode = ExtractionMode.BeforeDelimiter,
                PrimaryDelimiter = "-"
            };

            for (int i = 0; i < 1000; i++)
            {
                string input = $"PREFIX-{i}-SUFFIX";
                string result = TextExtractor.Extract(input, opts);
                Assert.AreEqual("PREFIX", result);
            }
        }

        [TestMethod]
        public void RepeatedExecution_DuplicateFinder_DeterministicAndIdempotent()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { "A", 1 },
                { "B", 2 },
                { "A", 1 },
                { "C", 3 }
            };

            var opts = new DuplicateDetectionOptions
            {
                KeepStrategy = DuplicateKeepStrategy.KeepFirst,
                MatchCase = false
            };

            for (int i = 0; i < 100; i++)
            {
                bool[] duplicates = DuplicateFinder.IdentifyDuplicateRows(matrix, opts);
                Assert.AreEqual(4, duplicates.Length);
                Assert.IsFalse(duplicates[0]);
                Assert.IsFalse(duplicates[1]);
                Assert.IsTrue(duplicates[2]);
                Assert.IsFalse(duplicates[3]);
            }
        }

        [TestMethod]
        public void AdvancedFilter_HandlesNullsBlanksAndEmptyCriteriaGracefully()
        {
            object?[,] matrix = new object?[4, 2]
            {
                { "Name", "Score" },
                { null, null },
                { "", 100 },
                { "ValidName", 200 }
            };

            var options = new AdvancedFilterOptions
            {
                Conjunction = FilterConjunction.And,
                HasHeaders = true,
                Conditions = new List<FilterCondition>
                {
                    new FilterCondition
                    {
                        ColumnIndex = 0,
                        Operator = FilterOperator.IsNotBlank,
                        ComparisonValue = ""
                    }
                }
            };

            var result = AdvancedFilter.Execute(matrix, options);
            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.MatchCount);
            Assert.AreEqual("ValidName", result.MatchingRows[0][0]);

            var exported = result.ExportToMatrix();
            Assert.AreEqual(2, exported.GetLength(0)); // Header + 1 matched row
        }

        [TestMethod]
        public void RangeMerger_HandlesEmptyAndSingleRowRangesGracefully()
        {
            object?[,] r1 = new object?[1, 2]
            {
                { "Name", "Role" }
            };
            object?[,] r2 = new object?[2, 2]
            {
                { "Dev", "Lead" },
                { "QA", "Tester" }
            };

            var options = new RangeMergeOptions
            {
                Mode = MergeMode.AppendRows,
                HasHeaders = false
            };

            var result = RangeMerger.Merge(r1, r2, options);
            Assert.IsNotNull(result);
            Assert.AreEqual(3, result.TotalRows);
            Assert.AreEqual(2, result.TotalColumns);
            Assert.IsNotNull(result.MergedMatrix);
        }

        [TestMethod]
        public void DataSplitter_SanitizeSheetName_HandlesAllExcelRestrictions()
        {
            // Test forbidden characters
            string dirty = "Data/Sheet:2026*Q3?[Draft]";
            string clean = DataSplitter.SanitizeSheetName(dirty);
            Assert.IsFalse(clean.Contains("/"));
            Assert.IsFalse(clean.Contains("\\"));
            Assert.IsFalse(clean.Contains(":"));
            Assert.IsFalse(clean.Contains("*"));
            Assert.IsFalse(clean.Contains("?"));
            Assert.IsFalse(clean.Contains("["));
            Assert.IsFalse(clean.Contains("]"));

            // Test 31 character limit
            string longName = "A Very Long Worksheet Name That Definitely Exceeds Thirty One Characters In Total";
            string truncated = DataSplitter.SanitizeSheetName(longName);
            Assert.IsTrue(truncated.Length <= 31);

            // Test reserved name History
            string reserved = DataSplitter.SanitizeSheetName("History");
            Assert.AreNotEqual("History", reserved);

            // Test empty or whitespace name
            string empty = DataSplitter.SanitizeSheetName("   ");
            Assert.AreEqual("Sheet", empty);
        }
    }
}
