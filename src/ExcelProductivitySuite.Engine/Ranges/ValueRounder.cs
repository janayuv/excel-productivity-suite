using System;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum RoundingMode
    {
        Standard,          // Math.Round away from zero
        RoundUpCeiling,    // Always round up toward positive infinity
        RoundDownFloor,    // Always round down toward negative infinity
        Truncate,          // Strip decimal places toward zero
        NearestMultiple    // Round to nearest multiple (e.g. 0.05, 0.5, 5, 10, 50)
    }

    public class ValueRoundOptions
    {
        public RoundingMode Mode { get; set; } = RoundingMode.Standard;
        public int DecimalPlaces { get; set; } = 2;
        public double Multiple { get; set; } = 5.0;
        public bool SkipNonNumeric { get; set; } = true;
    }

    /// <summary>
    /// Performs in-place rounding of numeric values across ranges without inserting formulas.
    /// </summary>
    public static class ValueRounder
    {
        public static object? Round(object? input, ValueRoundOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (input == null) return null;

            if (!TryGetNumeric(input, out double num))
            {
                return options.SkipNonNumeric ? input : null;
            }

            int decimals = Math.Max(0, options.DecimalPlaces);

            switch (options.Mode)
            {
                case RoundingMode.Standard:
                    return Math.Round(num, decimals, MidpointRounding.AwayFromZero);

                case RoundingMode.RoundUpCeiling:
                    double factorUp = Math.Pow(10, decimals);
                    return Math.Ceiling(num * factorUp) / factorUp;

                case RoundingMode.RoundDownFloor:
                    double factorDown = Math.Pow(10, decimals);
                    return Math.Floor(num * factorDown) / factorDown;

                case RoundingMode.Truncate:
                    double factorTrunc = Math.Pow(10, decimals);
                    return Math.Truncate(num * factorTrunc) / factorTrunc;

                case RoundingMode.NearestMultiple:
                    if (options.Multiple <= 0) return num;
                    return Math.Round(num / options.Multiple, MidpointRounding.AwayFromZero) * options.Multiple;

                default:
                    return Math.Round(num, decimals, MidpointRounding.AwayFromZero);
            }
        }

        public static object?[,] RoundRange(object?[,] source, ValueRoundOptions options, out int affectedCount)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            object?[,] result = (object?[,])source.Clone();
            affectedCount = 0;

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    object? val = source[r, c];
                    if (val != null && TryGetNumeric(val, out _))
                    {
                        result[r, c] = Round(val, options);
                        affectedCount++;
                    }
                }
            }

            return result;
        }

        private static bool TryGetNumeric(object value, out double number)
        {
            if (value is double d) { number = d; return true; }
            if (value is int i) { number = i; return true; }
            if (value is long l) { number = l; return true; }
            if (value is decimal m) { number = (double)m; return true; }
            if (value is float f) { number = f; return true; }
            if (value is short s) { number = s; return true; }

            if (value is string str && double.TryParse(str.Trim(), out double parsed))
            {
                number = parsed;
                return true;
            }

            number = 0;
            return false;
        }
    }
}
