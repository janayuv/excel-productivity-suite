using System;
using System.Globalization;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum MathOperationType
    {
        Add,
        Subtract,
        Multiply,
        Divide,
        Power,
        PercentChange
    }

    public class MathOperationOptions
    {
        public bool RoundResults { get; set; } = false;
        public int DecimalPlaces { get; set; } = 2;
        public bool TreatTextAsNumber { get; set; } = true;
        public bool SkipBlankCells { get; set; } = true;
    }

    public class MathOperationResult
    {
        public object?[,] Matrix { get; }
        public int ProcessedCount { get; }
        public int SkippedCount { get; }
        public int ErrorCount { get; }

        public MathOperationResult(object?[,] matrix, int processed, int skipped, int errors)
        {
            Matrix = matrix;
            ProcessedCount = processed;
            SkippedCount = skipped;
            ErrorCount = errors;
        }
    }

    /// <summary>
    /// Engine for applying arithmetic operations directly to cell values in bulk without formula overhead.
    /// </summary>
    public static class RangeMathOperations
    {
        public static MathOperationResult Apply(
            object?[,] matrix,
            MathOperationType operation,
            double operand,
            MathOperationOptions? options = null)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            options ??= new MathOperationOptions();

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rows = rEnd - rStart + 1;
            int cols = cEnd - cStart + 1;

            object?[,] result = new object?[rows, cols];
            int processed = 0;
            int skipped = 0;
            int errors = 0;

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    int outR = r - rStart;
                    int outC = c - cStart;
                    object? cell = matrix[r, c];

                    if (cell == null || (cell is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        if (options.SkipBlankCells)
                        {
                            result[outR, outC] = cell;
                            skipped++;
                            continue;
                        }
                    }

                    if (TryExtractNumericValue(cell, options.TreatTextAsNumber, out double numericVal))
                    {
                        try
                        {
                            double computed;
                            switch (operation)
                            {
                                case MathOperationType.Add:
                                    computed = numericVal + operand;
                                    break;
                                case MathOperationType.Subtract:
                                    computed = numericVal - operand;
                                    break;
                                case MathOperationType.Multiply:
                                    computed = numericVal * operand;
                                    break;
                                case MathOperationType.Divide:
                                    if (Math.Abs(operand) < double.Epsilon)
                                    {
                                        result[outR, outC] = "#DIV/0!";
                                        errors++;
                                        continue;
                                    }
                                    computed = numericVal / operand;
                                    break;
                                case MathOperationType.Power:
                                    computed = Math.Pow(numericVal, operand);
                                    break;
                                case MathOperationType.PercentChange:
                                    computed = numericVal * (1.0 + (operand / 100.0));
                                    break;
                                default:
                                    computed = numericVal;
                                    break;
                            }

                            if (options.RoundResults)
                            {
                                computed = Math.Round(computed, Math.Max(0, Math.Min(15, options.DecimalPlaces)), MidpointRounding.AwayFromZero);
                            }

                            result[outR, outC] = computed;
                            processed++;
                        }
                        catch
                        {
                            result[outR, outC] = cell;
                            errors++;
                        }
                    }
                    else
                    {
                        result[outR, outC] = cell;
                        skipped++;
                    }
                }
            }

            return new MathOperationResult(result, processed, skipped, errors);
        }

        private static bool TryExtractNumericValue(object? cell, bool treatTextAsNumber, out double value)
        {
            value = 0.0;
            if (cell == null) return false;

            if (cell is double d) { value = d; return true; }
            if (cell is int i) { value = i; return true; }
            if (cell is long l) { value = l; return true; }
            if (cell is float f) { value = f; return true; }
            if (cell is decimal dec) { value = (double)dec; return true; }
            if (cell is short s) { value = s; return true; }

            if (treatTextAsNumber && cell is string str)
            {
                str = str.Trim();
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
                {
                    value = parsed;
                    return true;
                }
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.CurrentCulture, out parsed))
                {
                    value = parsed;
                    return true;
                }
            }

            return false;
        }
    }
}
