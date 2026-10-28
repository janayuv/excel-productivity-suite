using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    public class DistinctExtractOptions
    {
        public bool HasHeaders { get; set; } = true;
        public int[]? KeyColumns { get; set; } = null; // 0-based column indices; if null, all columns
        public bool MatchCase { get; set; } = false;
    }

    /// <summary>
    /// Pure computational engine for extracting unique / distinct records from a dataset without mutating the source.
    /// </summary>
    public static class DistinctExtractor
    {
        public static object?[,] ExtractDistinct(
            object?[,] matrix,
            DistinctExtractOptions options,
            out int uniqueCount,
            out int duplicateCount)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            uniqueCount = 0;
            duplicateCount = 0;

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int totalRows = rEnd - rStart + 1;
            int totalCols = cEnd - cStart + 1;

            if (totalRows <= 0 || totalCols <= 0) return new object?[0, 0];

            int dataStart = options.HasHeaders ? rStart + 1 : rStart;
            var comparer = options.MatchCase ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            var seenKeys = new HashSet<string>(comparer);
            var retainedRowIndices = new List<int>();

            for (int r = dataStart; r <= rEnd; r++)
            {
                string rowKey = BuildRowKey(matrix, r, cStart, cEnd, options.KeyColumns);
                if (seenKeys.Add(rowKey))
                {
                    retainedRowIndices.Add(r);
                    uniqueCount++;
                }
                else
                {
                    duplicateCount++;
                }
            }

            int finalRows = (options.HasHeaders ? 1 : 0) + retainedRowIndices.Count;
            object?[,] result = new object?[finalRows, totalCols];
            int destRow = 0;

            if (options.HasHeaders)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    result[0, c - cStart] = matrix[rStart, c];
                }
                destRow = 1;
            }

            foreach (int sourceRow in retainedRowIndices)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    result[destRow, c - cStart] = matrix[sourceRow, c];
                }
                destRow++;
            }

            return result;
        }

        private static string BuildRowKey(object?[,] matrix, int row, int cStart, int cEnd, int[]? keyCols)
        {
            if (keyCols != null && keyCols.Length > 0)
            {
                var parts = new string[keyCols.Length];
                for (int i = 0; i < keyCols.Length; i++)
                {
                    int col = cStart + keyCols[i];
                    parts[i] = matrix[row, col]?.ToString()?.Trim() ?? "";
                }
                return string.Join("||", parts);
            }
            else
            {
                int count = cEnd - cStart + 1;
                var parts = new string[count];
                for (int c = cStart; c <= cEnd; c++)
                {
                    parts[c - cStart] = matrix[row, c]?.ToString()?.Trim() ?? "";
                }
                return string.Join("||", parts);
            }
        }
    }
}
