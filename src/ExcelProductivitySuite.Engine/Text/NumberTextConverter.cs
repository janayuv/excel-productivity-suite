using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Text
{
    public class NumberTextConversionOptions
    {
        public bool AllowCurrencySymbols { get; set; } = true;
        public bool AllowParenthesesForNegative { get; set; } = true;
        public bool AllowPercentage { get; set; } = true;
        public bool SkipNonNumeric { get; set; } = true;
        public bool TrimWhitespace { get; set; } = true;
    }

    public static class NumberTextConverter
    {
        private static readonly Regex CurrencySymbolsRegex = new Regex(@"[\$€£¥₹\s\u00A0]", RegexOptions.Compiled);

        public static object? ConvertToNumber(object? input, NumberTextConversionOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (input == null) return null;

            // If it's already a numeric type, return as is
            if (input is double || input is int || input is long || input is decimal || input is float || input is short)
            {
                return input;
            }

            string raw = input.ToString() ?? "";
            if (options.TrimWhitespace)
            {
                raw = raw.Trim();
            }

            if (string.IsNullOrEmpty(raw))
            {
                return options.SkipNonNumeric ? (object)raw : null;
            }

            string clean = raw;

            // Handle percentage (e.g. "25%" or "25.5 %")
            bool isPercent = false;
            if (options.AllowPercentage && clean.EndsWith("%"))
            {
                isPercent = true;
                clean = clean.Substring(0, clean.Length - 1).Trim();
            }

            // Handle accounting parentheses: "(1,234.50)" -> "-1,234.50"
            bool isNegative = false;
            if (options.AllowParenthesesForNegative && clean.StartsWith("(") && clean.EndsWith(")"))
            {
                isNegative = true;
                clean = clean.Substring(1, clean.Length - 2).Trim();
            }

            // Strip currency symbols
            if (options.AllowCurrencySymbols)
            {
                clean = CurrencySymbolsRegex.Replace(clean, "");
            }

            // Try standard invariant parsing first (e.g. "1,234.56" or "1234.56")
            NumberStyles styles = NumberStyles.Number | NumberStyles.AllowLeadingSign;
            if (double.TryParse(clean, styles, CultureInfo.InvariantCulture, out double result))
            {
                if (isNegative && result > 0) result = -result;
                if (isPercent) result = result / 100.0;
                return result;
            }

            // Try European culture where dot is thousand separator and comma is decimal (e.g. "1.234,56")
            var europeanCulture = CultureInfo.GetCultureInfo("de-DE");
            if (double.TryParse(clean, styles, europeanCulture, out double euResult))
            {
                if (isNegative && euResult > 0) euResult = -euResult;
                if (isPercent) euResult = euResult / 100.0;
                return euResult;
            }

            // Try CurrentCulture
            if (double.TryParse(clean, styles, CultureInfo.CurrentCulture, out double curResult))
            {
                if (isNegative && curResult > 0) curResult = -curResult;
                if (isPercent) curResult = curResult / 100.0;
                return curResult;
            }

            // If unparseable, return original string or null
            return options.SkipNonNumeric ? (object)raw : null;
        }
    }
}
