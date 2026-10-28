using System;
using System.Text;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum SpellerMode
    {
        Currency,
        PureNumber
    }

    public enum SpellerCurrency
    {
        USD, // Dollars & Cents
        EUR, // Euros & Cents
        GBP, // Pounds & Pence
        INR, // Rupees & Paise
        Generic // Units & Hundredths
    }

    public enum SpellerCase
    {
        TitleCase,
        UpperCase,
        LowerCase,
        SentenceCase
    }

    public class NumberSpellerOptions
    {
        public SpellerMode Mode { get; set; } = SpellerMode.Currency;
        public SpellerCurrency Currency { get; set; } = SpellerCurrency.USD;
        public SpellerCase LetterCase { get; set; } = SpellerCase.TitleCase;
        public bool IncludeCentsIfZero { get; set; } = true;
    }

    public static class NumberSpeller
    {
        private static readonly string[] Ones = new[]
        {
            "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
            "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"
        };

        private static readonly string[] Tens = new[]
        {
            "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"
        };

        public static string Spell(double value, NumberSpellerOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            bool isNegative = value < 0;
            double absValue = Math.Abs(value);

            long integerPart = (long)Math.Floor(absValue);
            int decimalPart = (int)Math.Round((absValue - integerPart) * 100);

            // Handle rounding edge case (e.g. 0.999 -> 1.00)
            if (decimalPart >= 100)
            {
                integerPart += 1;
                decimalPart = 0;
            }

            string integerWords = ConvertIntegerToWords(integerPart);
            string result;

            if (options.Mode == SpellerMode.Currency)
            {
                GetCurrencyNames(options.Currency, integerPart, decimalPart,
                    out string majorUnit, out string minorUnit);

                var sb = new StringBuilder();
                if (isNegative) sb.Append("Negative ");

                sb.Append(integerWords);
                sb.Append(" ");
                sb.Append(majorUnit);

                if (decimalPart > 0 || options.IncludeCentsIfZero)
                {
                    sb.Append(" and ");
                    sb.Append(ConvertIntegerToWords(decimalPart));
                    sb.Append(" ");
                    sb.Append(minorUnit);
                }

                result = sb.ToString();
            }
            else // PureNumber
            {
                var sb = new StringBuilder();
                if (isNegative) sb.Append("Negative ");

                sb.Append(integerWords);

                if (decimalPart > 0)
                {
                    sb.Append(" Point ");
                    int d1 = decimalPart / 10;
                    int d2 = decimalPart % 10;
                    sb.Append(Ones[d1]);
                    sb.Append(" ");
                    sb.Append(Ones[d2]);
                }

                result = sb.ToString();
            }

            return ApplyCase(result.Trim(), options.LetterCase);
        }

        private static string ConvertIntegerToWords(long number)
        {
            if (number == 0) return Ones[0];

            if (number < 0)
            {
                return "Negative " + ConvertIntegerToWords(Math.Abs(number));
            }

            var parts = new StringBuilder();

            if (number >= 1000000000000L) // Trillion
            {
                parts.Append(ConvertIntegerToWords(number / 1000000000000L) + " Trillion ");
                number %= 1000000000000L;
            }

            if (number >= 1000000000L) // Billion
            {
                parts.Append(ConvertIntegerToWords(number / 1000000000L) + " Billion ");
                number %= 1000000000L;
            }

            if (number >= 1000000L) // Million
            {
                parts.Append(ConvertIntegerToWords(number / 1000000L) + " Million ");
                number %= 1000000L;
            }

            if (number >= 1000L) // Thousand
            {
                parts.Append(ConvertIntegerToWords(number / 1000L) + " Thousand ");
                number %= 1000L;
            }

            if (number >= 100L) // Hundred
            {
                parts.Append(ConvertIntegerToWords(number / 100L) + " Hundred ");
                number %= 100L;
            }

            if (number > 0)
            {
                if (number < 20)
                {
                    parts.Append(Ones[number]);
                }
                else
                {
                    parts.Append(Tens[number / 10]);
                    if ((number % 10) > 0)
                    {
                        parts.Append("-" + Ones[number % 10]);
                    }
                }
            }

            return parts.ToString().Trim();
        }

        private static void GetCurrencyNames(SpellerCurrency currency, long major, int minor, out string majorName, out string minorName)
        {
            switch (currency)
            {
                case SpellerCurrency.USD:
                    majorName = major == 1 ? "Dollar" : "Dollars";
                    minorName = minor == 1 ? "Cent" : "Cents";
                    break;

                case SpellerCurrency.EUR:
                    majorName = major == 1 ? "Euro" : "Euros";
                    minorName = minor == 1 ? "Cent" : "Cents";
                    break;

                case SpellerCurrency.GBP:
                    majorName = major == 1 ? "Pound" : "Pounds";
                    minorName = minor == 1 ? "Penny" : "Pence";
                    break;

                case SpellerCurrency.INR:
                    majorName = major == 1 ? "Rupee" : "Rupees";
                    minorName = minor == 1 ? "Paisa" : "Paise";
                    break;

                case SpellerCurrency.Generic:
                default:
                    majorName = major == 1 ? "Unit" : "Units";
                    minorName = minor == 1 ? "Hundredth" : "Hundredths";
                    break;
            }
        }

        private static string ApplyCase(string text, SpellerCase letterCase)
        {
            switch (letterCase)
            {
                case SpellerCase.UpperCase:
                    return text.ToUpperInvariant();

                case SpellerCase.LowerCase:
                    return text.ToLowerInvariant();

                case SpellerCase.SentenceCase:
                    if (text.Length <= 1) return text.ToUpperInvariant();
                    return char.ToUpperInvariant(text[0]) + text.Substring(1).ToLowerInvariant();

                case SpellerCase.TitleCase:
                default:
                    return text;
            }
        }
    }
}
