using System;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum SignChangeMode
    {
        InvertSign,            // Positive -> Negative, Negative -> Positive
        MakeAllPositive,       // Absolute value |x|
        MakeAllNegative,       // Negative -|x|
        ReverseNegativeOnly,   // Change negative to positive, positive unchanged
        ReversePositiveOnly    // Change positive to negative, negative unchanged
    }

    public class SignChangeOptions
    {
        public SignChangeMode Mode { get; set; } = SignChangeMode.InvertSign;
        public bool SkipNonNumeric { get; set; } = true;
    }

    /// <summary>
    /// Inverts or alters the mathematical sign (+/-) of numeric cells in a range.
    /// </summary>
    public static class SignChanger
    {
        public static object? ChangeSign(object? input, SignChangeOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (input == null) return null;

            if (!TryGetNumeric(input, out double num))
            {
                return options.SkipNonNumeric ? input : null;
            }

            switch (options.Mode)
            {
                case SignChangeMode.InvertSign:
                    num = -num;
                    break;

                case SignChangeMode.MakeAllPositive:
                    num = Math.Abs(num);
                    break;

                case SignChangeMode.MakeAllNegative:
                    num = -Math.Abs(num);
                    break;

                case SignChangeMode.ReverseNegativeOnly:
                    if (num < 0) num = -num;
                    break;

                case SignChangeMode.ReversePositiveOnly:
                    if (num > 0) num = -num;
                    break;
            }

            // Preserve integer representation if original was integral
            if (input is int || input is long || input is short)
            {
                return (long)num;
            }

            return num;
        }

        public static object?[,] ChangeSignRange(object?[,] source, SignChangeOptions options, out int affectedCount)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            object?[,] result = (object?[,])source.Clone();
            affectedCount = 0;

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    object? val = source[r, c];
                    if (val != null && TryGetNumeric(val, out _))
                    {
                        result[r, c] = ChangeSign(val, options);
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
