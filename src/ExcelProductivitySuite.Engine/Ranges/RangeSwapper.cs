using System;

namespace ExcelProductivitySuite.Engine.Ranges
{
    /// <summary>
    /// Swaps data between two ranges of equal or arbitrary dimensions in managed memory.
    /// </summary>
    public static class RangeSwapper
    {
        public static (object?[,] SwappedA, object?[,] SwappedB) Swap(
            object?[,] rangeA,
            object?[,] rangeB,
            bool requireEqualDimensions = false)
        {
            if (rangeA == null) throw new ArgumentNullException(nameof(rangeA));
            if (rangeB == null) throw new ArgumentNullException(nameof(rangeB));

            int rCountA = rangeA.GetUpperBound(0) - rangeA.GetLowerBound(0) + 1;
            int cCountA = rangeA.GetUpperBound(1) - rangeA.GetLowerBound(1) + 1;

            int rCountB = rangeB.GetUpperBound(0) - rangeB.GetLowerBound(0) + 1;
            int cCountB = rangeB.GetUpperBound(1) - rangeB.GetLowerBound(1) + 1;

            if (requireEqualDimensions && (rCountA != rCountB || cCountA != cCountB))
            {
                throw new InvalidOperationException(
                    $"Ranges have different dimensions: Range 1 is {rCountA}x{cCountA}, Range 2 is {rCountB}x{cCountB}.");
            }

            // SwappedA receives rangeB, SwappedB receives rangeA
            return ((object?[,])rangeB.Clone(), (object?[,])rangeA.Clone());
        }
    }
}
