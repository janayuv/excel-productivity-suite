using System;
using System.Collections.Generic;
using System.Diagnostics;
using ExcelProductivitySuite.Engine.Analysis;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.Engine.Files;
using ExcelProductivitySuite.Engine.Formulas;
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

        [TestMethod]
        public void Benchmark_TextReverser_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = $"Row_{r}-Col_{c}-PayloadData";
                }
            }

            var options = new TextReverseOptions
            {
                Mode = TextReverseMode.ByCharacter
            };

            var sw = Stopwatch.StartNew();
            var result = TextReverser.ReverseRange(matrix, options);
            sw.Stop();

            Assert.AreEqual(10000, result.GetLength(0));
            Assert.AreEqual(10, result.GetLength(1));
            Assert.IsNotNull(result[0, 0]);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"TextReverser benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_TextCombiner_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = $"Part{c}";
                }
            }

            var options = new CombineOptions
            {
                Direction = CombineDirection.AcrossColumns,
                Delimiter = " | "
            };

            var sw = Stopwatch.StartNew();
            var result = TextCombiner.CombineMatrix(matrix, options);
            sw.Stop();

            Assert.AreEqual(10000, result.GetLength(0));
            Assert.AreEqual(1, result.GetLength(1));
            Assert.IsNotNull(result[0, 0]);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"TextCombiner benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_RangeTransposer_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = r * 10 + c;
                }
            }

            var sw = Stopwatch.StartNew();
            var transposed = RangeTransposer.Transpose(matrix);
            sw.Stop();

            Assert.AreEqual(10, transposed.GetLength(0));
            Assert.AreEqual(10000, transposed.GetLength(1));
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"RangeTransposer benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_RangeRandomizer_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = $"Cell_{r}_{c}";
                }
            }

            var options = new RandomSortOptions
            {
                Mode = RandomSortMode.EntireRows,
                HasHeaderRow = true
            };

            var sw = Stopwatch.StartNew();
            var shuffled = RangeRandomizer.Randomize(matrix, options, new Random(42));
            sw.Stop();

            Assert.AreEqual(10000, shuffled.GetLength(0));
            Assert.AreEqual(10, shuffled.GetLength(1));
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"RangeRandomizer benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_ValueRounder_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = (r * 1.234567) + (c * 0.987654);
                }
            }

            var options = new ValueRoundOptions
            {
                Mode = RoundingMode.Standard,
                DecimalPlaces = 2
            };

            var sw = Stopwatch.StartNew();
            var rounded = ValueRounder.RoundRange(matrix, options, out int affected);
            sw.Stop();

            Assert.AreEqual(100000, affected);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"ValueRounder benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_DuplicateAnalyzer_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    // Introduce duplicates every 5 rows
                    matrix[r, c] = $"Item_{r % 2000}_{c}";
                }
            }

            var options = new DuplicateAnalysisOptions
            {
                MatchRule = DuplicateMatchRule.DuplicatesExcludingFirst,
                Scope = DuplicateAnalysisScope.CellByCell,
                MatchCase = false
            };

            var sw = Stopwatch.StartNew();
            var result = DuplicateAnalyzer.Analyze(matrix, 1, 1, options);
            sw.Stop();

            Assert.AreEqual(100000, result.TotalCellsAnalyzed);
            Assert.IsTrue(result.MatchingCells.Count > 0);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"DuplicateAnalyzer benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_MultiSheetFindReplace_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = $"Sample string value {r}-{c} with target token in text";
                }
            }

            var options = new FindReplaceOptions
            {
                SearchText = "target token",
                ReplaceText = "replacement token",
                MatchCase = true
            };

            var sw = Stopwatch.StartNew();
            var matches = MultiSheetFindReplace.FindMatches("Sheet1", 1, 1, matrix, options);
            sw.Stop();

            Assert.AreEqual(100000, matches.Count);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"MultiSheetFindReplace benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_SheetCombiner_100K_Cells_CompletesUnder500ms()
        {
            int sheetCount = 5;
            int rowsPerSheet = 2000;
            int cols = 10;

            var sheets = new List<SheetDataInfo>();
            for (int s = 0; s < sheetCount; s++)
            {
                object?[,] data = new object?[rowsPerSheet, cols];
                for (int r = 0; r < rowsPerSheet; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        data[r, c] = (r == 0) ? $"Header_{c}" : $"Data_{s}_{r}_{c}";
                    }
                }
                sheets.Add(new SheetDataInfo($"Sheet_{s}", data));
            }

            var options = new SheetCombineOptions
            {
                HasHeaders = true,
                KeepFirstSheetHeaderOnly = true,
                AddSourceSheetColumn = true,
                BlankRowsBetweenSheets = 0
            };

            var sw = Stopwatch.StartNew();
            var combined = SheetCombiner.Combine(sheets, options, out int totalRows);
            sw.Stop();

            // 1 header row + 5 * 1999 data rows = 9996 rows x 11 cols (~110,000 cells)
            Assert.IsTrue(totalRows > 9900);
            Assert.AreEqual(11, combined.GetLength(1));
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"SheetCombiner benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_AdvancedSorter_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    // Generate variable length strings
                    matrix[r, c] = new string('X', (r * 7 + c * 3) % 50 + 1);
                }
            }

            var options = new AdvancedSortOptions
            {
                SortColumnIndex = 0,
                HasHeaderRow = false,
                Criterion = AdvancedSortCriterion.TextLengthAscending
            };

            var sw = Stopwatch.StartNew();
            var sorted = AdvancedSorter.Sort(matrix, options);
            sw.Stop();

            Assert.AreEqual(10000, sorted.GetLength(0));
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"AdvancedSorter benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_TableMerger_100K_Cells_CompletesUnder500ms()
        {
            int pRows = 10000;
            int pCols = 5;
            object?[,] primary = new object?[pRows, pCols];

            int lRows = 5000;
            int lCols = 5;
            object?[,] lookup = new object?[lRows, lCols];

            for (int r = 0; r < pRows; r++)
            {
                primary[r, 0] = $"ID_{r}";
                for (int c = 1; c < pCols; c++) primary[r, c] = $"ValP_{r}_{c}";
            }

            for (int r = 0; r < lRows; r++)
            {
                lookup[r, 0] = $"ID_{r * 2}"; // match half the rows
                for (int c = 1; c < lCols; c++) lookup[r, c] = $"ValL_{r}_{c}";
            }

            var options = new TableMergeOptions
            {
                PrimaryKeyColumn = 0,
                LookupKeyColumn = 0,
                HasHeaders = false,
                AppendUnmatchedRows = false
            };

            var sw = Stopwatch.StartNew();
            var merged = TableMerger.Merge(primary, lookup, options, out int matched, out int unmatched);
            sw.Stop();

            Assert.AreEqual(5000, matched);
            Assert.AreEqual(10000, merged.GetLength(0));
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"TableMerger benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_RangeMathOperations_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = (double)(r * 10 + c);
                }
            }

            var options = new MathOperationOptions
            {
                RoundResults = true,
                DecimalPlaces = 2,
                TreatTextAsNumber = true,
                SkipBlankCells = true
            };

            var sw = Stopwatch.StartNew();
            var result = RangeMathOperations.Apply(matrix, MathOperationType.Multiply, 1.15, options);
            sw.Stop();

            Assert.AreEqual(100000, result.ProcessedCount);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"RangeMathOperations benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_FormulaReferenceConverter_10K_Formulas_CompletesUnder500ms()
        {
            int rows = 2000;
            int cols = 5;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = $"=SUM(A{r + 1}:B{r + 10}) + C${r + 5} * $D$4";
                }
            }

            var sw = Stopwatch.StartNew();
            var (resMatrix, modified) = FormulaReferenceConverter.ConvertMatrix(matrix, ReferenceConversionMode.ToAbsolute);
            sw.Stop();

            Assert.AreEqual(10000, modified);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"FormulaReferenceConverter benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_CsvTextImporter_100K_Cells_CompletesUnder500ms()
        {
            var sb = new System.Text.StringBuilder();
            for (int r = 0; r < 10000; r++)
            {
                sb.AppendLine($"User_{r},Company_{r},12345,99.95,\"Description with, comma\",2026-01-01,Active,Cat_{r % 10},Ref_{r},Standard");
            }
            string csvContent = sb.ToString();

            var sw = Stopwatch.StartNew();
            var matrix = CsvTextImporter.ParseText(csvContent);
            sw.Stop();

            Assert.AreEqual(10000, matrix.GetLength(0));
            Assert.AreEqual(10, matrix.GetLength(1));
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"CsvTextImporter benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_RangeExporter_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = $"Cell_{r}_{c}";
                }
            }

            var options = new RangeExportOptions { Format = ExportFileFormat.Csv, QuoteAllFields = false };

            var sw = Stopwatch.StartNew();
            string csv = RangeExporter.ExportToString(matrix, options);
            sw.Stop();

            Assert.IsTrue(csv.Length > 100000);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"RangeExporter benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }

        [TestMethod]
        public void Benchmark_QuickStatisticsEngine_100K_Cells_CompletesUnder500ms()
        {
            int rows = 10000;
            int cols = 10;
            object?[,] matrix = new object?[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    matrix[r, c] = (double)(r * 10 + c);
                }
            }

            var sw = Stopwatch.StartNew();
            var stats = QuickStatisticsEngine.Calculate(matrix);
            sw.Stop();

            Assert.AreEqual(100000, stats.NumericCells);
            Assert.IsTrue(stats.Sum > 0);
            Assert.IsTrue(sw.ElapsedMilliseconds < 500, $"QuickStatisticsEngine benchmark took {sw.ElapsedMilliseconds} ms, target is < 500 ms");
        }
    }
}
