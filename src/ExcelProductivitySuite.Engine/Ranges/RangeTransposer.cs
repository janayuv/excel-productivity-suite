using System;

namespace ExcelProductivitySuite.Engine.Ranges
{
    /// <summary>
    /// Transposes rows to columns and columns to rows for 2D cell matrices.
    /// Handles both 0-based and 1-based (Excel COM) arrays while preserving cell types.
    /// </summary>
    public static class RangeTransposer
    {
        public static object?[,] Transpose(object?[,] source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            // Transposed matrix swaps dimensions: new rows = source cols, new cols = source rows
            object?[,] result = new object?[colCount, rowCount];

            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    result[c, r] = source[rStart + r, cStart + c];
                }
            }

            return result;
        }
    }
}
