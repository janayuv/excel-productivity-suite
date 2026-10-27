using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Engine.Ranges;

namespace ExcelProductivitySuite.Engine.Worksheets
{
    public class SheetComparisonReport
    {
        public string SheetName { get; set; } = "";
        public int RowsInA { get; set; }
        public int ColsInA { get; set; }
        public int RowsInB { get; set; }
        public int ColsInB { get; set; }
        public List<CellDifference> Discrepancies { get; } = new List<CellDifference>();
        public int DiscrepancyCount => Discrepancies.Count;
    }

    public class WorkbookCompareOptions
    {
        public bool MatchCase { get; set; } = false;
        public bool TrimWhitespace { get; set; } = true;
        public bool IgnoreEmptyCells { get; set; } = true;
    }

    public class WorkbookComparisonResult
    {
        public string WorkbookAName { get; set; } = "Workbook A";
        public string WorkbookBName { get; set; } = "Workbook B";

        public List<string> SheetsOnlyInA { get; } = new List<string>();
        public List<string> SheetsOnlyInB { get; } = new List<string>();
        public List<SheetComparisonReport> SheetReports { get; } = new List<SheetComparisonReport>();

        public int TotalDiscrepancies
        {
            get
            {
                int count = 0;
                foreach (var report in SheetReports) count += report.DiscrepancyCount;
                return count;
            }
        }

        public object?[,] BuildSummaryMatrix()
        {
            var rows = new List<object?[]>();

            // Title / Metadata rows
            rows.Add(new object?[] { "Workbook Comparison Summary Report", "", "", "" });
            rows.Add(new object?[] { "Workbook A:", WorkbookAName, "", "" });
            rows.Add(new object?[] { "Workbook B:", WorkbookBName, "", "" });
            rows.Add(new object?[] { "Sheets Only in A:", string.Join(", ", SheetsOnlyInA), "", "" });
            rows.Add(new object?[] { "Sheets Only in B:", string.Join(", ", SheetsOnlyInB), "", "" });
            rows.Add(new object?[] { "Total Cell Discrepancies:", TotalDiscrepancies, "", "" });
            rows.Add(new object?[] { "", "", "", "" }); // Empty spacer

            // Table Header
            rows.Add(new object?[] { "Worksheet", "Cell Location", $"{WorkbookAName} Value", $"{WorkbookBName} Value", "Discrepancy Type" });

            foreach (var report in SheetReports)
            {
                foreach (var diff in report.Discrepancies)
                {
                    rows.Add(new object?[]
                    {
                        report.SheetName,
                        $"R{diff.Row}C{diff.Column}",
                        diff.SourceValue ?? "(Empty)",
                        diff.TargetValue ?? "(Empty)",
                        diff.DifferenceType.ToString()
                    });
                }
            }

            int outRows = rows.Count;
            int outCols = 5;
            object?[,] matrix = new object?[outRows, outCols];

            for (int r = 0; r < outRows; r++)
            {
                var rValues = rows[r];
                for (int c = 0; c < Math.Min(outCols, rValues.Length); c++)
                {
                    matrix[r, c] = rValues[c];
                }
            }

            return matrix;
        }
    }

    public static class WorkbookComparer
    {
        public static WorkbookComparisonResult Compare(
            string workbookAName,
            Dictionary<string, object?[,]> workbookASheets,
            string workbookBName,
            Dictionary<string, object?[,]> workbookBSheets,
            WorkbookCompareOptions options)
        {
            if (workbookASheets == null) throw new ArgumentNullException(nameof(workbookASheets));
            if (workbookBSheets == null) throw new ArgumentNullException(nameof(workbookBSheets));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var result = new WorkbookComparisonResult
            {
                WorkbookAName = workbookAName,
                WorkbookBName = workbookBName
            };

            var aKeys = new HashSet<string>(workbookASheets.Keys, StringComparer.OrdinalIgnoreCase);
            var bKeys = new HashSet<string>(workbookBSheets.Keys, StringComparer.OrdinalIgnoreCase);

            // Sheets in A not in B
            foreach (var k in aKeys)
            {
                if (!bKeys.Contains(k)) result.SheetsOnlyInA.Add(k);
            }

            // Sheets in B not in A
            foreach (var k in bKeys)
            {
                if (!aKeys.Contains(k)) result.SheetsOnlyInB.Add(k);
            }

            // Common sheets: compare cell-by-cell using RangeComparer
            var rangeOptions = new RangeCompareOptions
            {
                Mode = RangeCompareMode.CellByCell,
                MatchCase = options.MatchCase,
                TrimWhitespace = options.TrimWhitespace,
                IgnoreEmptyCells = options.IgnoreEmptyCells
            };

            foreach (var sheetName in aKeys)
            {
                if (bKeys.Contains(sheetName))
                {
                    var matrixA = workbookASheets[sheetName];
                    var matrixB = workbookBSheets[sheetName];

                    var rangeResult = RangeComparer.Compare(matrixA, matrixB, rangeOptions);

                    var report = new SheetComparisonReport
                    {
                        SheetName = sheetName,
                        RowsInA = matrixA.GetLength(0),
                        ColsInA = matrixA.GetLength(1),
                        RowsInB = matrixB.GetLength(0),
                        ColsInB = matrixB.GetLength(1)
                    };

                    report.Discrepancies.AddRange(rangeResult.Differences);
                    result.SheetReports.Add(report);
                }
            }

            return result;
        }
    }
}
