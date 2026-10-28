using System;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum FlipDirection
    {
        Horizontal, // Columns reversed left-to-right
        Vertical    // Rows reversed top-to-bottom
    }

    /// <summary>
    /// Flips 2D ranges horizontally (column order reversal) or vertically (row order reversal).
    /// </summary>
    public static class RangeFlipper
    {
        public static object?[,] FlipHorizontal(object?[,] source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            object?[,] result = new object?[rowCount, colCount];

            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    // Target col c gets source col from the right: (colCount - 1 - c)
                    result[r, c] = source[rStart + r, cStart + (colCount - 1 - c)];
                }
            }

            return result;
        }

        public static object?[,] FlipVertical(object?[,] source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            object?[,] result = new object?[rowCount, colCount];

            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    // Target row r gets source row from bottom: (rowCount - 1 - r)
                    result[r, c] = source[rStart + (rowCount - 1 - r), cStart + c];
                }
            }

            return result;
        }

        public static object?[,] Flip(object?[,] source, FlipDirection direction)
        {
            return direction == FlipDirection.Horizontal
                ? FlipHorizontal(source)
                : FlipVertical(source);
        }
    }
}
