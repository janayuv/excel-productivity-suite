using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Engine.Worksheets;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class WorkbookComparerTests
    {
        [TestMethod]
        public void Compare_DetectsMissingSheetsAndCellDifferences()
        {
            var wbA = new Dictionary<string, object?[,]>
            {
                {
                    "Summary",
                    new object?[2, 2]
                    {
                        { "Metric", "Value" },
                        { "Revenue", 1000 }
                    }
                },
                {
                    "OldArchive",
                    new object?[1, 1] { { "ArchiveData" } }
                }
            };

            var wbB = new Dictionary<string, object?[,]>
            {
                {
                    "Summary",
                    new object?[2, 2]
                    {
                        { "Metric", "Value" },
                        { "Revenue", 1250 } // Mismatched value
                    }
                },
                {
                    "NewForecast",
                    new object?[1, 1] { { "ForecastData" } }
                }
            };

            var options = new WorkbookCompareOptions
            {
                MatchCase = false,
                TrimWhitespace = true
            };

            var result = WorkbookComparer.Compare("Book1.xlsx", wbA, "Book2.xlsx", wbB, options);

            // Sheet discrepancies
            Assert.AreEqual(1, result.SheetsOnlyInA.Count);
            Assert.AreEqual("OldArchive", result.SheetsOnlyInA[0]);

            Assert.AreEqual(1, result.SheetsOnlyInB.Count);
            Assert.AreEqual("NewForecast", result.SheetsOnlyInB[0]);

            // Cell discrepancies
            Assert.AreEqual(1, result.SheetReports.Count);
            Assert.AreEqual("Summary", result.SheetReports[0].SheetName);
            Assert.AreEqual(1, result.SheetReports[0].DiscrepancyCount);
            Assert.AreEqual("1000", result.SheetReports[0].Discrepancies[0].SourceValue);
            Assert.AreEqual("1250", result.SheetReports[0].Discrepancies[0].TargetValue);

            // Summary matrix export
            var summaryMatrix = result.BuildSummaryMatrix();
            Assert.IsTrue(summaryMatrix.GetLength(0) >= 8);
        }
    }
}
