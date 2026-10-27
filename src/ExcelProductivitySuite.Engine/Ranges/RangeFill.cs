using System;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum FillBlankDirection
    {
        Down,
        Up,
        Right,
        Left,
        FixedValue
    }

    public class FillBlankOptions
    {
        public FillBlankDirection Direction { get; set; } = FillBlankDirection.Down;
        public object? FixedValue { get; set; } = "";
    }

    public static class RangeFill
    {
        /// <summary>
        /// Fills blank cells in a 2D array based on the specified direction.
        /// Handles 1-based (Excel COM) or 0-based indexing seamlessly.
        /// </summary>
        public static object?[,] FillBlanks(object?[,] source, FillBlankOptions options, out int filledCount)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rowStart = source.GetLowerBound(0);
            int rowEnd = source.GetUpperBound(0);
            int colStart = source.GetLowerBound(1);
            int colEnd = source.GetUpperBound(1);

            int rows = rowEnd - rowStart + 1;
            int cols = colEnd - colStart + 1;

            var result = (object?[,])source.Clone();
            filledCount = 0;

            switch (options.Direction)
            {
                case FillBlankDirection.Down:
                    for (int c = colStart; c <= colEnd; c++)
                    {
                        object? lastVal = null;
                        for (int r = rowStart; r <= rowEnd; r++)
                        {
                            if (IsEmpty(result[r, c]))
                            {
                                if (lastVal != null)
                                {
                                    result[r, c] = lastVal;
                                    filledCount++;
                                }
                            }
                            else
                            {
                                lastVal = result[r, c];
                            }
                        }
                    }
                    break;

                case FillBlankDirection.Up:
                    for (int c = colStart; c <= colEnd; c++)
                    {
                        object? lastVal = null;
                        for (int r = rowEnd; r >= rowStart; r--)
                        {
                            if (IsEmpty(result[r, c]))
                            {
                                if (lastVal != null)
                                {
                                    result[r, c] = lastVal;
                                    filledCount++;
                                }
                            }
                            else
                            {
                                lastVal = result[r, c];
                            }
                        }
                    }
                    break;

                case FillBlankDirection.Right:
                    for (int r = rowStart; r <= rowEnd; r++)
                    {
                        object? lastVal = null;
                        for (int c = colStart; c <= colEnd; c++)
                        {
                            if (IsEmpty(result[r, c]))
                            {
                                if (lastVal != null)
                                {
                                    result[r, c] = lastVal;
                                    filledCount++;
                                }
                            }
                            else
                            {
                                lastVal = result[r, c];
                            }
                        }
                    }
                    break;

                case FillBlankDirection.Left:
                    for (int r = rowStart; r <= rowEnd; r++)
                    {
                        object? lastVal = null;
                        for (int c = colEnd; c >= colStart; c--)
                        {
                            if (IsEmpty(result[r, c]))
                            {
                                if (lastVal != null)
                                {
                                    result[r, c] = lastVal;
                                    filledCount++;
                                }
                            }
                            else
                            {
                                lastVal = result[r, c];
                            }
                        }
                    }
                    break;

                case FillBlankDirection.FixedValue:
                    for (int r = rowStart; r <= rowEnd; r++)
                    {
                        for (int c = colStart; c <= colEnd; c++)
                        {
                            if (IsEmpty(result[r, c]))
                            {
                                result[r, c] = options.FixedValue;
                                filledCount++;
                            }
                        }
                    }
                    break;
            }

            return result;
        }

        private static bool IsEmpty(object? val)
        {
            if (val == null) return true;
            if (val is string s && string.IsNullOrWhiteSpace(s)) return true;
            return false;
        }
    }
}
