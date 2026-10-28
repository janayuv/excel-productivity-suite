using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelProductivitySuite.Engine.Formatting
{
    public enum NumberFormatCategory
    {
        Currency,
        Accounting,
        Percentage,
        DecimalNumber,
        Integer,
        Scientific,
        CompactUnits,
        Custom
    }

    public enum NegativeNumberStyle
    {
        StandardMinus,       // -1,234.56
        Parentheses,         // (1,234.56)
        RedMinus,            // [Red]-1,234.56
        RedParentheses       // [Red](1,234.56)
    }

    public class NumberFormatPresetItem
    {
        public string Id { get; }
        public string Name { get; }
        public NumberFormatCategory Category { get; }
        public string FormatString { get; }
        public string Description { get; }

        public NumberFormatPresetItem(string id, string name, NumberFormatCategory category, string formatString, string description)
        {
            Id = id;
            Name = name;
            Category = category;
            FormatString = formatString;
            Description = description;
        }

        public override string ToString() => Name;
    }

    public static class NumberFormatPresets
    {
        private static readonly List<NumberFormatPresetItem> _presets = new List<NumberFormatPresetItem>
        {
            // Currency
            new NumberFormatPresetItem("curr_usd", "USD Currency ($#,##0.00)", NumberFormatCategory.Currency, "$#,##0.00;($#,##0.00);\"-\"", "Standard US Dollar currency format with dash for zero"),
            new NumberFormatPresetItem("curr_usd_red", "USD Red Negative ($#,##0.00;[Red]...)", NumberFormatCategory.Currency, "$#,##0.00;[Red]-$#,##0.00;$0.00", "US Dollar format highlighting negative amounts in red"),
            new NumberFormatPresetItem("curr_eur", "Euro (€#,##0.00)", NumberFormatCategory.Currency, "€#,##0.00;[Red]-€#,##0.00;€0.00", "Euro format with two decimals and red negatives"),
            new NumberFormatPresetItem("curr_gbp", "British Pound (£#,##0.00)", NumberFormatCategory.Currency, "£#,##0.00;[Red]-£#,##0.00;£0.00", "British Pound format with two decimals"),
            new NumberFormatPresetItem("curr_inr", "Indian Rupee (₹#,##0.00)", NumberFormatCategory.Currency, "₹#,##0.00;[Red]-₹#,##0.00;₹0.00", "Indian Rupee standard format"),
            new NumberFormatPresetItem("curr_jpy", "Japanese Yen (¥#,##0)", NumberFormatCategory.Currency, "¥#,##0;[Red]-¥#,##0;¥0", "Japanese Yen without decimal places"),

            // Accounting
            new NumberFormatPresetItem("acct_usd", "USD Accounting Standard", NumberFormatCategory.Accounting, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)", "Aligned currency symbol and aligned parentheses for negative numbers"),
            new NumberFormatPresetItem("acct_eur", "Euro Accounting Standard", NumberFormatCategory.Accounting, "_(€* #,##0.00_);_(€* (#,##0.00);_(€* \"-\"??_);_(@_)", "Euro accounting format with margin-aligned symbols"),

            // Percentage
            new NumberFormatPresetItem("pct_2dec", "Percentage 2 Decimals (0.00%)", NumberFormatCategory.Percentage, "0.00%", "Percentage displayed with two decimal places"),
            new NumberFormatPresetItem("pct_1dec", "Percentage 1 Decimal (0.0%)", NumberFormatCategory.Percentage, "0.0%", "Percentage displayed with single decimal precision"),
            new NumberFormatPresetItem("pct_int", "Percentage Integer (0%)", NumberFormatCategory.Percentage, "0%", "Whole number percentage without decimals"),
            new NumberFormatPresetItem("pct_red", "Percentage Red Negative (0.00%;[Red]...)", NumberFormatCategory.Percentage, "0.00%;[Red]-0.00%;0.00%", "Percentage with negative values displayed in red"),

            // Decimal & Number
            new NumberFormatPresetItem("num_comma_2dec", "Standard Number (1,234.56)", NumberFormatCategory.DecimalNumber, "#,##0.00", "Comma thousands separator with two decimal places"),
            new NumberFormatPresetItem("num_comma_parens", "Number with Parentheses (1,234.56;(1,234.56))", NumberFormatCategory.DecimalNumber, "#,##0.00;(#,##0.00);\"-\"", "Accounting-style parentheses for negative values"),
            new NumberFormatPresetItem("num_comma_red", "Number with Red Negatives", NumberFormatCategory.DecimalNumber, "#,##0.00;[Red]-#,##0.00;0.00", "Red text for negative numbers"),
            new NumberFormatPresetItem("num_4dec", "Precision Number (4 Decimals)", NumberFormatCategory.DecimalNumber, "#,##0.0000", "High precision four decimal places"),

            // Integer
            new NumberFormatPresetItem("int_comma", "Integer with Separator (#,##0)", NumberFormatCategory.Integer, "#,##0", "Whole number with thousands separator"),
            new NumberFormatPresetItem("int_parens", "Integer Parentheses (#,##0;(#,##0))", NumberFormatCategory.Integer, "#,##0;(#,##0);0", "Whole number with parentheses for negative values"),

            // Scientific
            new NumberFormatPresetItem("sci_std", "Scientific Notation (0.00E+00)", NumberFormatCategory.Scientific, "0.00E+00", "Standard exponential notation with two decimals"),

            // Compact Units (Financial K / M / B)
            new NumberFormatPresetItem("unit_millions", "Compact Millions ($#,##0.0,,\" M\")", NumberFormatCategory.CompactUnits, "$#,##0.0,,\" M\"", "Scales figures down by millions with M suffix"),
            new NumberFormatPresetItem("unit_thousands", "Compact Thousands ($#,##0.0,\" K\")", NumberFormatCategory.CompactUnits, "$#,##0.0,\" K\"", "Scales figures down by thousands with K suffix")
        };

        public static IReadOnlyList<NumberFormatPresetItem> GetAllPresets() => _presets;

        public static string BuildFormat(
            NumberFormatCategory category,
            int decimals,
            bool useThousandSeparator,
            NegativeNumberStyle negativeStyle,
            string currencySymbol = "$")
        {
            decimals = Math.Max(0, Math.Min(6, decimals));

            // Decimal portion
            string decPart = decimals > 0 ? "." + new string('0', decimals) : "";

            // Base positive number pattern
            string baseNum = useThousandSeparator ? $"#,##0{decPart}" : $"0{decPart}";

            switch (category)
            {
                case NumberFormatCategory.Currency:
                {
                    string curr = string.IsNullOrWhiteSpace(currencySymbol) ? "$" : currencySymbol;
                    string pos = $"{curr}{baseNum}";
                    string neg = negativeStyle switch
                    {
                        NegativeNumberStyle.Parentheses => $"({curr}{baseNum})",
                        NegativeNumberStyle.RedMinus => $"[Red]-{curr}{baseNum}",
                        NegativeNumberStyle.RedParentheses => $"[Red]({curr}{baseNum})",
                        _ => $"-{curr}{baseNum}"
                    };
                    return $"{pos};{neg};{curr}0{decPart}";
                }

                case NumberFormatCategory.Accounting:
                {
                    string curr = string.IsNullOrWhiteSpace(currencySymbol) ? "$" : currencySymbol;
                    return $"_({curr}* {baseNum}_);_({curr}* ({baseNum});_({curr}* \"-\"??_);_(@_)";
                }

                case NumberFormatCategory.Percentage:
                {
                    string pos = $"{baseNum}%";
                    string neg = negativeStyle switch
                    {
                        NegativeNumberStyle.Parentheses => $"({baseNum}%)",
                        NegativeNumberStyle.RedMinus => $"[Red]-{baseNum}%",
                        NegativeNumberStyle.RedParentheses => $"[Red]({baseNum}%)",
                        _ => $"-{baseNum}%"
                    };
                    return $"{pos};{neg};0{decPart}%";
                }

                case NumberFormatCategory.Integer:
                {
                    string pos = useThousandSeparator ? "#,##0" : "0";
                    string neg = negativeStyle switch
                    {
                        NegativeNumberStyle.Parentheses => $"({pos})",
                        NegativeNumberStyle.RedMinus => $"[Red]-{pos}",
                        NegativeNumberStyle.RedParentheses => $"[Red]({pos})",
                        _ => $"-{pos}"
                    };
                    return $"{pos};{neg};0";
                }

                case NumberFormatCategory.Scientific:
                {
                    string sciDec = decimals > 0 ? "." + new string('0', decimals) : "";
                    return $"0{sciDec}E+00";
                }

                case NumberFormatCategory.CompactUnits:
                {
                    string curr = string.IsNullOrWhiteSpace(currencySymbol) ? "$" : currencySymbol;
                    return $"{curr}#,##0.0,,\" M\"";
                }

                case NumberFormatCategory.DecimalNumber:
                default:
                {
                    string pos = baseNum;
                    string neg = negativeStyle switch
                    {
                        NegativeNumberStyle.Parentheses => $"({baseNum})",
                        NegativeNumberStyle.RedMinus => $"[Red]-{baseNum}",
                        NegativeNumberStyle.RedParentheses => $"[Red]({baseNum})",
                        _ => $"-{baseNum}"
                    };
                    return $"{pos};{neg};0{decPart}";
                }
            }
        }

        /// <summary>
        /// Generates a human-readable simulated preview string for a given numeric sample value.
        /// </summary>
        public static string FormatSamplePreview(double value, string formatString)
        {
            if (string.IsNullOrWhiteSpace(formatString)) return value.ToString(CultureInfo.InvariantCulture);

            try
            {
                // Clean Excel-specific color tokens like [Red], [Blue] and alignment tokens for .NET preview
                string netFmt = formatString
                    .Replace("[Red]", "")
                    .Replace("[Green]", "")
                    .Replace("[Blue]", "")
                    .Replace("_($*", "$")
                    .Replace("_(€*", "€")
                    .Replace("_(£*", "£")
                    .Replace("_(*", "")
                    .Replace("_)", "")
                    .Replace("_-", "")
                    .Replace("@", "");

                // Excel format strings have up to 4 sections: Positive;Negative;Zero;Text
                string[] sections = netFmt.Split(';');
                if (sections.Length > 0)
                {
                    if (value > 0)
                    {
                        return value.ToString(sections[0].Trim(), CultureInfo.InvariantCulture);
                    }
                    else if (value < 0 && sections.Length > 1)
                    {
                        double absVal = Math.Abs(value);
                        string negPattern = sections[1].Trim();
                        if (negPattern.StartsWith("-"))
                        {
                            return "-" + absVal.ToString(negPattern.Substring(1).Trim(), CultureInfo.InvariantCulture);
                        }
                        return absVal.ToString(negPattern, CultureInfo.InvariantCulture);
                    }
                    else if (value == 0 && sections.Length > 2)
                    {
                        string zeroPattern = sections[2].Trim().Trim('"');
                        return zeroPattern;
                    }
                }

                return value.ToString("N2", CultureInfo.InvariantCulture);
            }
            catch
            {
                return value.ToString("N2", CultureInfo.InvariantCulture);
            }
        }
    }
}
