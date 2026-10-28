using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum RandomSortMode
    {
        EntireRows,          // Shuffles rows preserving column alignment (records)
        EntireColumns,       // Shuffles column order
        CellsWithinColumns,  // Shuffles values independently within each column
        CellsWithinRows      // Shuffles values independently within each row
    }

    public class RandomSortOptions
    {
        public RandomSortMode Mode { get; set; } = RandomSortMode.EntireRows;
        public bool HasHeaderRow { get; set; } = false;
    }

    /// <summary>
    /// Randomly sorts or shuffles cell matrices using the Fisher-Yates shuffle algorithm.
    /// </summary>
    public static class RangeRandomizer
    {
        public static object?[,] Randomize(object?[,] source, RandomSortOptions options, Random? customRandom = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var rng = customRandom ?? new Random();

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            object?[,] result = (object?[,])source.Clone();

            int dataStartRow = (options.HasHeaderRow && rowCount > 1) ? rStart + 1 : rStart;
            int dataRowCount = rEnd - dataStartRow + 1;

            if (dataRowCount <= 1 && options.Mode == RandomSortMode.EntireRows)
            {
                return result;
            }

            switch (options.Mode)
            {
                case RandomSortMode.EntireRows:
                    // Create index map for rows
                    var rowIndices = new List<int>(dataRowCount);
                    for (int r = dataStartRow; r <= rEnd; r++) rowIndices.Add(r);

                    // Fisher-Yates shuffle on row indices
                    for (int i = rowIndices.Count - 1; i > 0; i--)
                    {
                        int j = rng.Next(i + 1);
                        int temp = rowIndices[i];
                        rowIndices[i] = rowIndices[j];
                        rowIndices[j] = temp;
                    }

                    // Copy shuffled rows
                    for (int i = 0; i < dataRowCount; i++)
                    {
                        int srcRow = rowIndices[i];
                        int dstRow = dataStartRow + i;
                        for (int c = cStart; c <= cEnd; c++)
                        {
                            result[dstRow, c] = source[srcRow, c];
                        }
                    }
                    break;

                case RandomSortMode.EntireColumns:
                    var colIndices = new List<int>(colCount);
                    for (int c = cStart; c <= cEnd; c++) colIndices.Add(c);

                    for (int i = colIndices.Count - 1; i > 0; i--)
                    {
                        int j = rng.Next(i + 1);
                        int temp = colIndices[i];
                        colIndices[i] = colIndices[j];
                        colIndices[j] = temp;
                    }

                    for (int i = 0; i < colCount; i++)
                    {
                        int srcCol = colIndices[i];
                        int dstCol = cStart + i;
                        for (int r = rStart; r <= rEnd; r++)
                        {
                            result[r, dstCol] = source[r, srcCol];
                        }
                    }
                    break;

                case RandomSortMode.CellsWithinColumns:
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        var cellValues = new List<object?>(dataRowCount);
                        for (int r = dataStartRow; r <= rEnd; r++) cellValues.Add(source[r, c]);

                        for (int i = cellValues.Count - 1; i > 0; i--)
                        {
                            int j = rng.Next(i + 1);
                            var temp = cellValues[i];
                            cellValues[i] = cellValues[j];
                            cellValues[j] = temp;
                        }

                        for (int i = 0; i < dataRowCount; i++)
                        {
                            result[dataStartRow + i, c] = cellValues[i];
                        }
                    }
                    break;

                case RandomSortMode.CellsWithinRows:
                    for (int r = dataStartRow; r <= rEnd; r++)
                    {
                        var cellValues = new List<object?>(colCount);
                        for (int c = cStart; c <= cEnd; c++) cellValues.Add(source[r, c]);

                        for (int i = cellValues.Count - 1; i > 0; i--)
                        {
                            int j = rng.Next(i + 1);
                            var temp = cellValues[i];
                            cellValues[i] = cellValues[j];
                            cellValues[j] = temp;
                        }

                        for (int i = 0; i < colCount; i++)
                        {
                            result[r, cStart + i] = cellValues[i];
                        }
                    }
                    break;
            }

            return result;
        }
    }
}
