using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    public enum DuplicateMatchRule
    {
        DuplicatesExcludingFirst, // 2nd, 3rd, ... occurrences (classic duplicate filter)
        AllDuplicates,            // Every occurrence of values that appear > 1 time
        UniqueValuesOnly          // Values that appear exactly 1 time
    }

    public enum DuplicateAnalysisScope
    {
        CellByCell,  // Evaluates every cell independently
        EntireRows   // Evaluates entire rows (or selected key columns)
    }

    public class DuplicateAnalysisOptions
    {
        public DuplicateMatchRule MatchRule { get; set; } = DuplicateMatchRule.DuplicatesExcludingFirst;
        public DuplicateAnalysisScope Scope { get; set; } = DuplicateAnalysisScope.CellByCell;
        public bool MatchCase { get; set; } = false;
        public bool HasHeaders { get; set; } = false;
        public int[]? KeyColumns { get; set; } = null; // 0-based column indices for row comparison
    }

    public class DuplicateAnalysisResult
    {
        public List<(int Row, int Column, object? Value)> MatchingCells { get; set; } = new();
        public int TotalCellsAnalyzed { get; set; }
        public int DuplicateGroupCount { get; set; }
        public int UniqueValueCount { get; set; }
    }

    /// <summary>
    /// Analyzes 2D ranges for duplicate and unique values with high performance O(N) hashing.
    /// </summary>
    public static class DuplicateAnalyzer
    {
        public static DuplicateAnalysisResult Analyze(object?[,] matrix, int startRow, int startCol, DuplicateAnalysisOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int totalRows = rEnd - rStart + 1;
            int totalCols = cEnd - cStart + 1;

            var result = new DuplicateAnalysisResult
            {
                TotalCellsAnalyzed = totalRows * totalCols
            };

            var comparer = options.MatchCase ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            // Map key -> list of cell coordinates
            var occurrences = new Dictionary<string, List<(int row, int col, object? val)>>(comparer);

            int dataStartRow = (options.HasHeaders && totalRows > 1) ? rStart + 1 : rStart;

            if (options.Scope == DuplicateAnalysisScope.CellByCell)
            {
                for (int r = dataStartRow; r <= rEnd; r++)
                {
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        object? val = matrix[r, c];
                        string str = val?.ToString()?.Trim() ?? "";
                        if (string.IsNullOrEmpty(str)) continue; // ignore blanks

                        int sheetRow = startRow + (r - rStart);
                        int sheetCol = startCol + (c - cStart);

                        if (!occurrences.TryGetValue(str, out var list))
                        {
                            list = new List<(int, int, object?)>();
                            occurrences[str] = list;
                        }
                        list.Add((sheetRow, sheetCol, val));
                    }
                }
            }
            else // EntireRows
            {
                for (int r = dataStartRow; r <= rEnd; r++)
                {
                    string rowKey = BuildRowKey(matrix, r, cStart, cEnd, options);
                    if (string.IsNullOrEmpty(rowKey)) continue;

                    int sheetRow = startRow + (r - rStart);

                    if (!occurrences.TryGetValue(rowKey, out var list))
                    {
                        list = new List<(int, int, object?)>();
                        occurrences[rowKey] = list;
                    }

                    // For rows, add all columns in the row as matching cells
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        int sheetCol = startCol + (c - cStart);
                        list.Add((sheetRow, sheetCol, matrix[r, c]));
                    }
                }
            }

            int dupeGroups = 0;
            int uniqueCount = 0;

            foreach (var kvp in occurrences)
            {
                var list = kvp.Value;
                int count = options.Scope == DuplicateAnalysisScope.EntireRows ? (list.Count / totalCols) : list.Count;

                if (count > 1)
                {
                    dupeGroups++;
                    if (options.MatchRule == DuplicateMatchRule.AllDuplicates)
                    {
                        result.MatchingCells.AddRange(list);
                    }
                    else if (options.MatchRule == DuplicateMatchRule.DuplicatesExcludingFirst)
                    {
                        if (options.Scope == DuplicateAnalysisScope.EntireRows)
                        {
                            // Skip the first row's cells
                            for (int i = totalCols; i < list.Count; i++)
                            {
                                result.MatchingCells.Add(list[i]);
                            }
                        }
                        else
                        {
                            // Skip first cell
                            for (int i = 1; i < list.Count; i++)
                            {
                                result.MatchingCells.Add(list[i]);
                            }
                        }
                    }
                }
                else
                {
                    uniqueCount++;
                    if (options.MatchRule == DuplicateMatchRule.UniqueValuesOnly)
                    {
                        result.MatchingCells.AddRange(list);
                    }
                }
            }

            result.DuplicateGroupCount = dupeGroups;
            result.UniqueValueCount = uniqueCount;
            return result;
        }

        private static string BuildRowKey(object?[,] matrix, int row, int colStart, int colEnd, DuplicateAnalysisOptions options)
        {
            if (options.KeyColumns != null && options.KeyColumns.Length > 0)
            {
                var parts = new string[options.KeyColumns.Length];
                for (int i = 0; i < options.KeyColumns.Length; i++)
                {
                    int c = colStart + options.KeyColumns[i];
                    if (c <= colEnd)
                    {
                        parts[i] = matrix[row, c]?.ToString()?.Trim() ?? "";
                    }
                }
                return string.Join("§|§", parts);
            }
            else
            {
                int cols = colEnd - colStart + 1;
                var parts = new string[cols];
                for (int c = colStart; c <= colEnd; c++)
                {
                    parts[c - colStart] = matrix[row, c]?.ToString()?.Trim() ?? "";
                }
                return string.Join("§|§", parts);
            }
        }
    }
}
