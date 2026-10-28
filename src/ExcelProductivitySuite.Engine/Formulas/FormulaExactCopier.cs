using System;

namespace ExcelProductivitySuite.Engine.Formulas
{
    /// <summary>
    /// Engine for copying formulas verbatim into target ranges without shifting relative cell references.
    /// Unlike standard Excel Copy/Paste, relative references like =A1+B1 remain exactly =A1+B1.
    /// </summary>
    public static class FormulaExactCopier
    {
        /// <summary>
        /// Duplicates a single formula verbatim across a target matrix of dimensions [rows, cols].
        /// </summary>
        public static object?[,] FillExactFormula(string formula, int rows, int cols)
        {
            if (rows <= 0) throw new ArgumentOutOfRangeException(nameof(rows));
            if (cols <= 0) throw new ArgumentOutOfRangeException(nameof(cols));

            object?[,] result = new object?[rows, cols];
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    result[r, c] = formula;
                }
            }
            return result;
        }

        /// <summary>
        /// Replicates a source formula matrix verbatim to target dimensions.
        /// If target dimensions exceed source dimensions, tiles or truncates the source pattern.
        /// </summary>
        public static object?[,] CopyExactMatrix(object?[,] sourceMatrix, int targetRows, int targetCols)
        {
            if (sourceMatrix == null) throw new ArgumentNullException(nameof(sourceMatrix));
            if (targetRows <= 0) throw new ArgumentOutOfRangeException(nameof(targetRows));
            if (targetCols <= 0) throw new ArgumentOutOfRangeException(nameof(targetCols));

            int sRowStart = sourceMatrix.GetLowerBound(0);
            int sRowEnd = sourceMatrix.GetUpperBound(0);
            int sColStart = sourceMatrix.GetLowerBound(1);
            int sColEnd = sourceMatrix.GetUpperBound(1);

            int sRows = sRowEnd - sRowStart + 1;
            int sCols = sColEnd - sColStart + 1;

            object?[,] target = new object?[targetRows, targetCols];

            for (int r = 0; r < targetRows; r++)
            {
                int srcR = sRowStart + (r % sRows);
                for (int c = 0; c < targetCols; c++)
                {
                    int srcC = sColStart + (c % sCols);
                    target[r, c] = sourceMatrix[srcR, srcC] ?? string.Empty;
                }
            }

            return target;
        }
    }
}
