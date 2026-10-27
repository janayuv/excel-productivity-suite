using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.Engine.Worksheets;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class PerformanceBenchmarkTests
    {
        [TestMethod]
        public void Benchmark_TextExtraction_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = $"User-{r}_{c}@subdomain.company.com";
                }
            }

            var options = new TextExtractionOptions
            {
                Mode = ExtractionMode.BeforeDelimiter,
                PrimaryDelimiter = "@"
            };

            var sw = Stopwatch.StartNew();

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int processed = 0;
            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    string str = matrix[r, c]?.ToString() ?? "";
                    string extracted = TextExtractor.Extract(str, options);
                    if (!string.IsNullOrEmpty(extracted)) processed++;
                }
            }

            sw.Stop();

            Assert.AreEqual(100000, processed);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"Benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_DuplicateDetection_100K_Cells_CompletesUnder1000ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                // Create intentional duplicates every 10 rows
                int keyId = r % 1000;
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = $"Val_{keyId}_{c}";
                }
            }

            var options = new DuplicateDetectionOptions
            {
                KeepStrategy = DuplicateKeepStrategy.KeepFirst,
                MatchCase = false
            };

            var sw = Stopwatch.StartNew();
            bool[] duplicates = DuplicateFinder.IdentifyDuplicateRows(matrix, options);
            sw.Stop();

            int dupCount = 0;
            foreach (bool d in duplicates) if (d) dupCount++;

            Assert.AreEqual(10000, duplicates.Length);
            Assert.AreEqual(9000, dupCount);
            Assert.IsTrue(sw.ElapsedMilliseconds < 1000, $"Duplicate benchmark took {sw.ElapsedMilliseconds} ms, target is < 1000 ms");
        }

        [TestMethod]
        public void Benchmark_FillBlanks_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    // 50% blank cells
                    matrix[r, c] = (r % 2 == 0) ? $"Header_{c}" : null;
                }
            }

            var options = new FillBlankOptions
            {
                Direction = FillBlankDirection.Down
            };

            var sw = Stopwatch.StartNew();
            var result = RangeFill.FillBlanks(matrix, options, out int filledCount);
            sw.Stop();

            Assert.AreEqual(50000, filledCount);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"FillBlanks benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_RangeComparer_100K_Cells_CompletesUnder1000ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] source = new object?[rows, cols];
            object?[,] target = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    source[r, c] = $"Item_{r}_{c}";
                    target[r, c] = (r % 100 == 0) ? $"Mismatch_{r}_{c}" : $"Item_{r}_{c}";
                }
            }

            var options = new RangeCompareOptions
            {
                Mode = RangeCompareMode.CellByCell,
                MatchCase = true
            };

            var sw = Stopwatch.StartNew();
            var result = RangeComparer.Compare(source, target, options);
            sw.Stop();

            Assert.AreEqual(100000, result.TotalCellsCompared);
            Assert.AreEqual(1000, result.TotalDifferences);
            Assert.IsTrue(sw.ElapsedMilliseconds < 1000, $"RangeComparer benchmark took {sw.ElapsedMilliseconds} ms, target is < 1000 ms");
        }

        [TestMethod]
        public void Benchmark_RangeRecordComparer_100K_Cells_CompletesUnder1000ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] source = new object?[rows, cols];
            object?[,] target = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    source[r, c] = (c == 0) ? $"ID_{r}" : $"Val_{r}_{c}";
                    target[r, c] = (c == 0) ? (r % 2 == 0 ? $"ID_{r}" : $"ID_T_{r}") : $"TVal_{r}_{c}";
                }
            }

            var options = new RecordCompareOptions
            {
                HasHeaders = true,
                SourceKeyColumns = new[] { 0 },
                TargetKeyColumns = new[] { 0 }
            };

            var sw = Stopwatch.StartNew();
            var result = RangeRecordComparer.CompareRecords(source, target, options);
            sw.Stop();

            Assert.IsTrue(result.TotalSourceRows > 0);
            Assert.IsTrue(sw.ElapsedMilliseconds < 1000, $"RecordComparer benchmark took {sw.ElapsedMilliseconds} ms, target is < 1000 ms");
        }

        [TestMethod]
        public void Benchmark_RangeMerger_100K_Cells_CompletesUnder500ms()
        {
            int rows = 5000;
            int cols = 10;
            object?[,] r1 = new object?[rows, cols];
            object?[,] r2 = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    r1[r, c] = (r == 0) ? $"Col_{c}" : $"Val1_{r}_{c}";
                    r2[r, c] = (r == 0) ? $"Col_{c}" : $"Val2_{r}_{c}";
                }
            }

            var options = new RangeMergeOptions
            {
                Mode = MergeMode.AppendRows,
                HasHeaders = true,
                AlignByHeaderNames = true,
                IncludeHeadersOnce = true
            };

            var sw = Stopwatch.StartNew();
            var result = RangeMerger.Merge(r1, r2, options);
            sw.Stop();

            Assert.AreEqual(9999, result.TotalRows); // 1 header + 4999 data + 4999 data
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"RangeMerger benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_DataSplitter_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    // 10 distinct departments
                    matrix[r, c] = (c == 0) ? $"Dept_{r % 10}" : $"Val_{r}_{c}";
                }
            }

            var options = new DataSplitOptions
            {
                SplitColumnIndex = 0,
                HasHeaders = true
            };

            var sw = Stopwatch.StartNew();
            var result = DataSplitter.Split(matrix, options);
            sw.Stop();

            Assert.AreEqual(10, result.DistinctGroupsCount);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"DataSplitter benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_AdvancedFilter_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = (c == 0) ? $"Region_{r % 5}" : (r * 10);
                }
            }

            var options = new AdvancedFilterOptions
            {
                Conjunction = FilterConjunction.And,
                HasHeaders = true,
                Conditions = new List<FilterCondition>
                {
                    new FilterCondition { ColumnIndex = 0, Operator = FilterOperator.Equals, ComparisonValue = "Region_1" },
                    new FilterCondition { ColumnIndex = 1, Operator = FilterOperator.GreaterThan, ComparisonValue = "1000" }
                }
            };

            var sw = Stopwatch.StartNew();
            var result = AdvancedFilter.Execute(matrix, options);
            sw.Stop();

            Assert.IsTrue(result.MatchCount > 0);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"AdvancedFilter benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }
    }
}
