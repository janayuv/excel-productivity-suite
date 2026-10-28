using System;

namespace ExcelProductivitySuite.Engine.Worksheets
{
    /// <summary>
    /// Evaluates 2D worksheet matrix data to verify whether a worksheet is completely blank.
    /// </summary>
    public static class BlankSheetDetector
    {
        public static bool IsMatrixBlank(object?[,]? matrix)
        {
            if (matrix == null) return true;

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rows = rEnd - rStart + 1;
            int cols = cEnd - cStart + 1;

            if (rows <= 0 || cols <= 0) return true;

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    var val = matrix[r, c];
                    if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                    {
                        return false;
                    }
                }
            }

            return true;
        }
    }
}
