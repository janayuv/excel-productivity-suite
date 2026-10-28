using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Abstractions;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public class UnmergeResult
    {
        public object?[,] ProcessedValues { get; }
        public int MergedAreasProcessed { get; }
        public int CellsFilled { get; }

        public UnmergeResult(object?[,] processedValues, int mergedAreasProcessed, int cellsFilled)
        {
            ProcessedValues = processedValues;
            MergedAreasProcessed = mergedAreasProcessed;
            CellsFilled = cellsFilled;
        }
    }

    public static class UnmergeAndFillEngine
    {
        /// <summary>
        /// Populates all cells in each merged area block with the primary composite value.
        /// </summary>
        public static UnmergeResult Process(
            object?[,] matrix,
            IReadOnlyList<MergedAreaInfo> mergedAreas,
            int selectionStartRow = 1,
            int selectionStartCol = 1)
        {
            if (matrix == null)
            {
                return new UnmergeResult(new object?[0, 0], 0, 0);
            }

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rows = rEnd - rStart + 1;
            int cols = cEnd - cStart + 1;
            var result = (object?[,])matrix.Clone();

            if (mergedAreas == null || mergedAreas.Count == 0)
            {
                return new UnmergeResult(result, 0, 0);
            }

            int areasProcessed = 0;
            int cellsFilled = 0;

            foreach (var area in mergedAreas)
            {
                object? fillVal = area.Value;
                if (fillVal == null) continue;

                areasProcessed++;

                // Map absolute sheet coordinates of merged area to local matrix coordinates
                int localStartR = (area.StartRow - selectionStartRow) + rStart;
                int localStartC = (area.StartCol - selectionStartCol) + cStart;

                for (int r = 0; r < area.RowCount; r++)
                {
                    int targetR = localStartR + r;
                    if (targetR < rStart || targetR > rEnd) continue;

                    for (int c = 0; c < area.ColCount; c++)
                    {
                        int targetC = localStartC + c;
                        if (targetC < cStart || targetC > cEnd) continue;

                        // Fill cell if currently null or empty (or not the top-left)
                        if (r > 0 || c > 0)
                        {
                            result[targetR, targetC] = fillVal;
                            cellsFilled++;
                        }
                    }
                }
            }

            return new UnmergeResult(result, areasProcessed, cellsFilled);
        }
    }
}
