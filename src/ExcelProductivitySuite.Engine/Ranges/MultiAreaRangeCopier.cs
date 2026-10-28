using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum MultiAreaCombineMode
    {
        StackVertically,    // Rows placed under rows
        PlaceSideBySide     // Columns placed next to columns
    }

    public class MultiAreaCopyOptions
    {
        public MultiAreaCombineMode Mode { get; set; } = MultiAreaCombineMode.StackVertically;
        public bool SkipBlankRows { get; set; } = false;
        public object? DefaultPadValue { get; set; } = null;
    }

    public static class MultiAreaRangeCopier
    {
        /// <summary>
        /// Combines multiple disjoint 2D matrices into a single contiguous 2D matrix.
        /// </summary>
        public static object?[,] Combine(IReadOnlyList<object?[,]> areas, MultiAreaCopyOptions? options = null)
        {
            options ??= new MultiAreaCopyOptions();

            if (areas == null || areas.Count == 0)
            {
                return new object?[0, 0];
            }

            if (options.Mode == MultiAreaCombineMode.StackVertically)
            {
                // Calculate max columns and total rows
                int maxCols = 0;
                int totalRows = 0;

                foreach (var area in areas)
                {
                    if (area == null) continue;
                    int rCount = area.GetUpperBound(0) - area.GetLowerBound(0) + 1;
                    int cCount = area.GetUpperBound(1) - area.GetLowerBound(1) + 1;
                    totalRows += rCount;
                    if (cCount > maxCols) maxCols = cCount;
                }

                if (totalRows == 0 || maxCols == 0) return new object?[0, 0];

                var result = new object?[totalRows, maxCols];
                int currentRow = 0;

                foreach (var area in areas)
                {
                    if (area == null) continue;
                    int rStart = area.GetLowerBound(0);
                    int rEnd = area.GetUpperBound(0);
                    int cStart = area.GetLowerBound(1);
                    int cEnd = area.GetUpperBound(1);
                    int cCount = cEnd - cStart + 1;

                    for (int r = rStart; r <= rEnd; r++)
                    {
                        for (int c = cStart; c <= cEnd; c++)
                        {
                            result[currentRow, c - cStart] = area[r, c];
                        }
                        // Pad any missing columns with default value
                        for (int c = cCount; c < maxCols; c++)
                        {
                            result[currentRow, c] = options.DefaultPadValue;
                        }
                        currentRow++;
                    }
                }

                return result;
            }
            else // PlaceSideBySide
            {
                int maxRows = 0;
                int totalCols = 0;

                foreach (var area in areas)
                {
                    if (area == null) continue;
                    int rCount = area.GetUpperBound(0) - area.GetLowerBound(0) + 1;
                    int cCount = area.GetUpperBound(1) - area.GetLowerBound(1) + 1;
                    totalCols += cCount;
                    if (rCount > maxRows) maxRows = rCount;
                }

                if (maxRows == 0 || totalCols == 0) return new object?[0, 0];

                var result = new object?[maxRows, totalCols];
                int currentCol = 0;

                foreach (var area in areas)
                {
                    if (area == null) continue;
                    int rStart = area.GetLowerBound(0);
                    int rEnd = area.GetUpperBound(0);
                    int cStart = area.GetLowerBound(1);
                    int cEnd = area.GetUpperBound(1);
                    int rCount = rEnd - rStart + 1;
                    int cCount = cEnd - cStart + 1;

                    for (int r = 0; r < maxRows; r++)
                    {
                        for (int c = cStart; c <= cEnd; c++)
                        {
                            if (r < rCount)
                            {
                                result[r, currentCol + (c - cStart)] = area[rStart + r, c];
                            }
                            else
                            {
                                result[r, currentCol + (c - cStart)] = options.DefaultPadValue;
                            }
                        }
                    }
                    currentCol += cCount;
                }

                return result;
            }
        }
    }
}
