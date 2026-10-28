using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelProductivitySuite.Engine.Navigation
{
    public class DateFormatOption
    {
        public string Name { get; }
        public string FormatString { get; }
        public string Example { get; }

        public DateFormatOption(string name, string formatString, DateTime sampleDate)
        {
            Name = name;
            FormatString = formatString;
            Example = sampleDate.ToString(formatString, CultureInfo.InvariantCulture);
        }

        public override string ToString() => $"{Name} ({Example})";
    }

    public static class DateFormatPresets
    {
        public static IReadOnlyList<DateFormatOption> GetStandardPresets(DateTime? sampleDate = null)
        {
            var date = sampleDate ?? DateTime.Now;
            return new List<DateFormatOption>
            {
                new DateFormatOption("ISO 8601 Date", "yyyy-MM-dd", date),
                new DateFormatOption("US Short Date", "MM/dd/yyyy", date),
                new DateFormatOption("European Short Date", "dd/MM/yyyy", date),
                new DateFormatOption("Long Date", "MMMM dd, yyyy", date),
                new DateFormatOption("Day & Date", "ddd, MMM dd, yyyy", date),
                new DateFormatOption("Standard Timestamp", "yyyy-MM-dd HH:mm:ss", date),
                new DateFormatOption("Year & Month", "yyyy-MM", date),
                new DateFormatOption("Full Month & Year", "MMMM yyyy", date)
            };
        }

        public static string Format(DateTime date, string formatPattern)
        {
            if (string.IsNullOrWhiteSpace(formatPattern))
            {
                formatPattern = "yyyy-MM-dd";
            }
            try
            {
                return date.ToString(formatPattern, CultureInfo.InvariantCulture);
            }
            catch
            {
                return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
        }
    }
}
