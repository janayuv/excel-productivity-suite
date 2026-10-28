using System;
using System.Globalization;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum DateOutputFormat
    {
        Iso8601,        // yyyy-MM-dd
        ShortDateUS,    // MM/dd/yyyy
        ShortDateEU,    // dd/MM/yyyy
        LongDate,       // MMMM dd, yyyy
        DateTimeObject  // DateTime struct for Excel serial dates
    }

    public enum DateCulturePreference
    {
        Auto,
        MonthDayYear, // US: MM/DD/YYYY
        DayMonthYear  // UK/EU: DD/MM/YYYY
    }

    public class DateTextConversionOptions
    {
        public DateOutputFormat OutputFormat { get; set; } = DateOutputFormat.Iso8601;
        public DateCulturePreference CulturePreference { get; set; } = DateCulturePreference.Auto;
        public bool SkipUnparseable { get; set; } = true;
    }

    public static class DateTextConverter
    {
        private static readonly string[] CommonDateFormats = new[]
        {
            "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd", "yyyyMMdd",
            "MM/dd/yyyy", "M/d/yyyy", "MM-dd-yyyy", "M-d-yyyy",
            "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "dd.MM.yyyy",
            "dd-MMM-yyyy", "dd MMM yyyy", "d MMM yyyy",
            "MMM dd, yyyy", "MMMM dd, yyyy", "dd MMMM yyyy",
            "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss", "MM/dd/yyyy HH:mm:ss"
        };

        public static object? ConvertToDate(object? input, DateTextConversionOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (input == null) return null;

            if (input is DateTime dt)
            {
                return FormatDate(dt, options.OutputFormat);
            }

            string raw = input.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(raw))
            {
                return options.SkipUnparseable ? (object)raw : null;
            }

            DateTime parsed = default;
            bool success = false;

            // 1. Try explicit formats based on culture preference
            if (options.CulturePreference == DateCulturePreference.DayMonthYear)
            {
                var euCulture = CultureInfo.GetCultureInfo("en-GB");
                success = DateTime.TryParse(raw, euCulture, DateTimeStyles.None, out parsed);
            }
            else if (options.CulturePreference == DateCulturePreference.MonthDayYear)
            {
                var usCulture = CultureInfo.GetCultureInfo("en-US");
                success = DateTime.TryParse(raw, usCulture, DateTimeStyles.None, out parsed);
            }

            // 2. Try common date formats array with InvariantCulture
            if (!success)
            {
                success = DateTime.TryParseExact(raw, CommonDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed);
            }

            // 3. Fallback to general InvariantCulture and CurrentCulture parsing
            if (!success)
            {
                success = DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed);
            }

            if (!success)
            {
                success = DateTime.TryParse(raw, CultureInfo.CurrentCulture, DateTimeStyles.None, out parsed);
            }

            if (success)
            {
                return FormatDate(parsed, options.OutputFormat);
            }

            return options.SkipUnparseable ? (object)raw : null;
        }

        private static object FormatDate(DateTime date, DateOutputFormat format)
        {
            switch (format)
            {
                case DateOutputFormat.Iso8601:
                    return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                case DateOutputFormat.ShortDateUS:
                    return date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

                case DateOutputFormat.ShortDateEU:
                    return date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

                case DateOutputFormat.LongDate:
                    return date.ToString("MMMM dd, yyyy", CultureInfo.InvariantCulture);

                case DateOutputFormat.DateTimeObject:
                    return date;

                default:
                    return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
        }
    }
}
