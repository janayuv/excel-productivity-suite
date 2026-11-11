using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Ranges
{
    /// <summary>
    /// Core calculation engine for scanning, grouping, and compacting completely empty rows.
    /// Powers the implementation of ASAP Utilities Tool #232 (Delete All Empty Rows).
    /// </summary>
    public static class DeleteEmptyRowsEngine
    {
        /// <summary>
        /// Determines whether an individual cell value is considered empty.
        /// </summary>
        public static bool IsEmptyValue(object? value, bool treatWhitespaceAsEmpty = true, bool treatEmptyFormulaAsEmpty = true)
        {
            if (value == null) return true;
            if (value is DBNull) return true;

            if (value is string s)
            {
                if (s.Length == 0) return true;
                if (treatWhitespaceAsEmpty && string.IsNullOrWhiteSpace(s)) return true;
                return false;
            }

            // Treat null-equivalent text representations from formulas (e.g. empty formula string)
            string strVal = value.ToString() ?? string.Empty;
            if (treatEmptyFormulaAsEmpty && strVal.Length == 0) return true;
            if (treatWhitespaceAsEmpty && string.IsNullOrWhiteSpace(strVal)) return true;

            return false;
        }

        /// <summary>
        /// Determines whether an entire row within a 2D matrix is completely empty.
        /// </summary>
        public static bool IsRowEmpty(
            object?[,] matrix,
            int rowIndex,
            bool treatWhitespaceAsEmpty = true,
            bool treatEmptyFormulaAsEmpty = true)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));

            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            for (int c = cStart; c <= cEnd; c++)
            {
                object? cellVal = matrix[rowIndex, c];
                if (!IsEmptyValue(cellVal, treatWhitespaceAsEmpty, treatEmptyFormulaAsEmpty))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Scans a 2D matrix and returns the 1-based worksheet row numbers that are completely empty.
        /// </summary>
        /// <param name="matrix">The cell values matrix.</param>
        /// <param name="startRowOffset">The 1-based sheet row index corresponding to the first matrix row.</param>
        /// <param name="treatWhitespaceAsEmpty">Whether to treat whitespace-only cells as empty.</param>
        /// <param name="treatEmptyFormulaAsEmpty">Whether to treat empty formula strings as empty.</param>
        /// <returns>A list of 1-based worksheet row indices that are empty.</returns>
        public static List<int> ScanEmptyRows(
            object?[,]? matrix,
            int startRowOffset = 1,
            bool treatWhitespaceAsEmpty = true,
            bool treatEmptyFormulaAsEmpty = true)
        {
            var emptyRows = new List<int>();
            if (matrix == null) return emptyRows;

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);

            for (int r = rStart; r <= rEnd; r++)
            {
                if (IsRowEmpty(matrix, r, treatWhitespaceAsEmpty, treatEmptyFormulaAsEmpty))
                {
                    int sheetRow = startRowOffset + (r - rStart);
                    emptyRows.Add(sheetRow);
                }
            }

            return emptyRows;
        }

        /// <summary>
        /// Groups consecutive row numbers into contiguous (startRow, endRow) intervals.
        /// E.g. [3, 4, 5, 8, 10, 11] => [(3, 5), (8, 8), (10, 11)].
        /// </summary>
        public static List<(int startRow, int endRow)> GroupContiguousRanges(IEnumerable<int>? rowIndices)
        {
            var ranges = new List<(int startRow, int endRow)>();
            if (rowIndices == null) return ranges;

            var sorted = rowIndices.Distinct().OrderBy(r => r).ToList();
            if (sorted.Count == 0) return ranges;

            int rangeStart = sorted[0];
            int rangeEnd = sorted[0];

            for (int i = 1; i < sorted.Count; i++)
            {
                int current = sorted[i];
                if (current == rangeEnd + 1)
                {
                    rangeEnd = current;
                }
                else
                {
                    ranges.Add((rangeStart, rangeEnd));
                    rangeStart = current;
                    rangeEnd = current;
                }
            }

            ranges.Add((rangeStart, rangeEnd));
            return ranges;
        }

        /// <summary>
        /// Builds comma-separated Excel range address chunks from contiguous row ranges.
        /// E.g. "3:5,8:8,10:11", bounded by maxCharsPerChunk to fit within Excel Range limits.
        /// Chunks are ordered from bottom-to-top (descending) so deleting them in order never shifts subsequent indices.
        /// </summary>
        public static List<string> BuildChunkedAddresses(
            List<(int startRow, int endRow)> contiguousRanges,
            int maxCharsPerChunk = 240,
            bool bottomToTop = true)
        {
            var chunks = new List<string>();
            if (contiguousRanges == null || contiguousRanges.Count == 0) return chunks;

            // Sort descending (bottom to top) for safe, order-independent deletion
            var ordered = bottomToTop
                ? contiguousRanges.OrderByDescending(r => r.startRow).ToList()
                : contiguousRanges.OrderBy(r => r.startRow).ToList();

            var sb = new StringBuilder();

            foreach (var (start, end) in ordered)
            {
                string segment = (start == end) ? $"{start}:{start}" : $"{start}:{end}";

                if (sb.Length == 0)
                {
                    sb.Append(segment);
                }
                else if (sb.Length + 1 + segment.Length <= maxCharsPerChunk)
                {
                    sb.Append(',').Append(segment);
                }
                else
                {
                    chunks.Add(sb.ToString());
                    sb.Clear();
                    sb.Append(segment);
                }
            }

            if (sb.Length > 0)
            {
                chunks.Add(sb.ToString());
            }

            return chunks;
        }

        /// <summary>
        /// Creates an in-memory compacted 2D array by stripping out all completely empty rows.
        /// </summary>
        public static object?[,] CompactMatrix(
            object?[,]? source,
            out int deletedCount,
            bool treatWhitespaceAsEmpty = true,
            bool treatEmptyFormulaAsEmpty = true)
        {
            if (source == null)
            {
                deletedCount = 0;
                return new object?[0, 0];
            }

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int totalRows = rEnd - rStart + 1;
            int totalCols = cEnd - cStart + 1;

            if (totalRows == 0 || totalCols == 0)
            {
                deletedCount = 0;
                return source;
            }

            var survivingRowIndices = new List<int>();
            for (int r = rStart; r <= rEnd; r++)
            {
                if (!IsRowEmpty(source, r, treatWhitespaceAsEmpty, treatEmptyFormulaAsEmpty))
                {
                    survivingRowIndices.Add(r);
                }
            }

            deletedCount = totalRows - survivingRowIndices.Count;
            if (survivingRowIndices.Count == 0)
            {
                return new object?[0, totalCols];
            }

            object?[,] result = new object?[survivingRowIndices.Count, totalCols];
            for (int i = 0; i < survivingRowIndices.Count; i++)
            {
                int srcRow = survivingRowIndices[i];
                for (int c = 0; c < totalCols; c++)
                {
                    result[i, c] = source[srcRow, cStart + c];
                }
            }

            return result;
        }

        /// <summary>
        /// Generates a comprehensive scan result for a worksheet matrix.
        /// </summary>
        public static EmptyRowsScanResult ScanWorksheetMatrix(
            string sheetName,
            object?[,]? matrix,
            int startRow = 1,
            int startCol = 1,
            bool treatWhitespaceAsEmpty = true,
            bool treatEmptyFormulaAsEmpty = true)
        {
            var result = new EmptyRowsScanResult
            {
                WorksheetName = sheetName ?? string.Empty,
                StartRow = startRow,
                StartCol = startCol
            };

            if (matrix == null)
            {
                result.TotalRowsScanned = 0;
                return result;
            }

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            result.TotalRowsScanned = rowCount;
            result.EndRow = startRow + rowCount - 1;
            result.EndCol = startCol + colCount - 1;

            var emptyIndices = ScanEmptyRows(matrix, startRow, treatWhitespaceAsEmpty, treatEmptyFormulaAsEmpty);
            result.EmptyRowIndices = emptyIndices;
            result.EmptyRowRanges = GroupContiguousRanges(emptyIndices);

            return result;
        }
    }
}
