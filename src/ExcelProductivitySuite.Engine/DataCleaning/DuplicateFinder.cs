using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    public enum DuplicateKeepStrategy
    {
        KeepFirst,
        KeepLast
    }

    public class DuplicateDetectionOptions
    {
        public DuplicateKeepStrategy KeepStrategy { get; set; } = DuplicateKeepStrategy.KeepFirst;
        public bool MatchCase { get; set; } = false;
        public bool CompareEntireRows { get; set; } = true;
        public bool HasHeaders { get; set; } = false;
        public int[]? KeyColumnIndices { get; set; } = null; // 0-based relative column offsets
    }

    public static class DuplicateFinder
    {
        /// <summary>
        /// Identifies duplicate rows in a 2D matrix. Returns a boolean array where true indicates the row is a duplicate to be removed/cleared.
        /// </summary>
        public static bool[] IdentifyDuplicateRows(object?[,] matrix, DuplicateDetectionOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rowStart = matrix.GetLowerBound(0);
            int rowEnd = matrix.GetUpperBound(0);
            int colStart = matrix.GetLowerBound(1);
            int colEnd = matrix.GetUpperBound(1);

            int totalRows = rowEnd - rowStart + 1;
            bool[] isDuplicate = new bool[totalRows];

            var comparer = options.MatchCase ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            var seenKeys = new HashSet<string>(comparer);

            int effectiveStart = options.HasHeaders ? rowStart + 1 : rowStart;

            if (options.KeepStrategy == DuplicateKeepStrategy.KeepFirst)
            {
                for (int r = effectiveStart; r <= rowEnd; r++)
                {
                    int rowIndex = r - rowStart;
                    string key = BuildRowKey(matrix, r, colStart, colEnd, options);

                    if (seenKeys.Contains(key))
                    {
                        isDuplicate[rowIndex] = true;
                    }
                    else
                    {
                        seenKeys.Add(key);
                    }
                }
            }
            else // KeepLast
            {
                for (int r = rowEnd; r >= effectiveStart; r--)
                {
                    int rowIndex = r - rowStart;
                    string key = BuildRowKey(matrix, r, colStart, colEnd, options);

                    if (seenKeys.Contains(key))
                    {
                        isDuplicate[rowIndex] = true;
                    }
                    else
                    {
                        seenKeys.Add(key);
                    }
                }
            }

            return isDuplicate;
        }

        private static string BuildRowKey(object?[,] matrix, int row, int colStart, int colEnd, DuplicateDetectionOptions options)
        {
            var parts = new List<string>();

            if (options.KeyColumnIndices != null && options.KeyColumnIndices.Length > 0)
            {
                foreach (int colOffset in options.KeyColumnIndices)
                {
                    int targetCol = colStart + colOffset;
                    if (targetCol <= colEnd)
                    {
                        parts.Add(matrix[row, targetCol]?.ToString() ?? "");
                    }
                }
            }
            else
            {
                for (int c = colStart; c <= colEnd; c++)
                {
                    parts.Add(matrix[row, c]?.ToString() ?? "");
                }
            }

            return string.Join("│", parts);
        }
    }
}
