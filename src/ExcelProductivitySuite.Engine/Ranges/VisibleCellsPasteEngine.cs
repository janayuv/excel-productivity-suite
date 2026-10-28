using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public class VisibleRowPasteMapping
    {
        public int SourceRowIndex { get; }
        public int TargetRowIndex { get; }
        public object?[] Values { get; }

        public VisibleRowPasteMapping(int sourceRow, int targetRow, object?[] values)
        {
            SourceRowIndex = sourceRow;
            TargetRowIndex = targetRow;
            Values = values;
        }
    }

    public static class VisibleCellsPasteEngine
    {
        /// <summary>
        /// Computes exact row-by-row mapping from source matrix to non-hidden visible rows.
        /// </summary>
        public static IReadOnlyList<VisibleRowPasteMapping> ComputeMapping(
            object?[,] sourceMatrix,
            IReadOnlyList<int> visibleRowIndices)
        {
            var mappings = new List<VisibleRowPasteMapping>();
            if (sourceMatrix == null || visibleRowIndices == null || visibleRowIndices.Count == 0)
            {
                return mappings;
            }

            int rStart = sourceMatrix.GetLowerBound(0);
            int rEnd = sourceMatrix.GetUpperBound(0);
            int cStart = sourceMatrix.GetLowerBound(1);
            int cEnd = sourceMatrix.GetUpperBound(1);
            int colCount = cEnd - cStart + 1;

            int srcRowCount = rEnd - rStart + 1;
            int countToMap = Math.Min(srcRowCount, visibleRowIndices.Count);

            for (int i = 0; i < countToMap; i++)
            {
                int srcR = rStart + i;
                int targetR = visibleRowIndices[i];

                var rowVals = new object?[colCount];
                for (int c = 0; c < colCount; c++)
                {
                    rowVals[c] = sourceMatrix[srcR, cStart + c];
                }

                mappings.Add(new VisibleRowPasteMapping(srcR, targetR, rowVals));
            }

            return mappings;
        }
    }
}
