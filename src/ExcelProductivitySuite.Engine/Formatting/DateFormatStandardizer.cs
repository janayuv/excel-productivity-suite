using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelProductivitySuite.Engine.Formatting
{
    public class DateFormatStandardItem
    {
        public string Id { get; }
        public string Name { get; }
        public string ExcelFormat { get; }
        public string SampleOutput { get; }
        public string Description { get; }

        public DateFormatStandardItem(string id, string name, string excelFormat, string description, DateTime? sample = null)
        {
            Id = id;
            Name = name;
            ExcelFormat = excelFormat;
            Description = description;
            DateTime dt = sample ?? DateTime.Now;
            SampleOutput = FormatSample(dt, excelFormat);
        }

        private static string FormatSample(DateTime dt, string fmt)
        {
            try
            {
                // Translate Excel date format tokens to .NET format tokens for display
                string netFmt = fmt
                    .Replace("yyyy", "YYYY")
                    .Replace("yy", "YY")
                    .Replace("mmmm", "MMMM")
                    .Replace("mmm", "MMM")
                    .Replace("mm", "MM")
                    .Replace("m", "M")
                    .Replace("dddd", "DDDD")
                    .Replace("ddd", "DDD")
                    .Replace("dd", "dd")
                    .Replace("hh", "HH")
                    .Replace("ss", "ss")
                    .Replace("YYYY", "yyyy")
                    .Replace("YY", "yy")
                    .Replace("DDDD", "dddd")
                    .Replace("DDD", "ddd");

                return dt.ToString(netFmt, CultureInfo.InvariantCulture);
            }
            catch
            {
                return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
        }

        public override string ToString() => $"{Name} ({SampleOutput})";
    }

    public class DateStandardizeResult
    {
        public object?[,] ProcessedValues { get; }
        public int TotalCells { get; }
        public int DateCellsFound { get; }
        public int TextDatesConverted { get; }
        public int UnchangedCells { get; }
        public string TargetFormat { get; }

        public DateStandardizeResult(
            object?[,] processedValues,
            int totalCells,
            int dateCellsFound,
            int textDatesConverted,
            int unchangedCells,
            string targetFormat)
        {
            ProcessedValues = processedValues;
            TotalCells = totalCells;
            DateCellsFound = dateCellsFound;
            TextDatesConverted = textDatesConverted;
            UnchangedCells = unchangedCells;
            TargetFormat = targetFormat;
        }
    }

    public static class DateFormatStandardizer
    {
        private static readonly string[] RecognizedParseFormats = new[]
        {
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyy.MM.dd",
            "dd/MM/yyyy",
            "dd-MM-yyyy",
            "dd.MM.yyyy",
            "MM/dd/yyyy",
            "MM-dd-yyyy",
            "yyyy-MM-dd HH:mm:ss",
            "yyyy/MM/dd HH:mm:ss",
            "dd/MM/yyyy HH:mm:ss",
            "MM/dd/yyyy HH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss",
            "dd-MMM-yyyy",
            "dd MMM yyyy",
            "MMMM dd, yyyy",
            "MMM dd, yyyy",
            "d/M/yyyy",
            "M/d/yyyy"
        };

        private static readonly List<DateFormatStandardItem> _standards = new List<DateFormatStandardItem>
        {
            new DateFormatStandardItem("iso_date", "ISO 8601 Date", "yyyy-mm-dd", "Standard international date format (YYYY-MM-DD)"),
            new DateFormatStandardItem("iso_datetime", "ISO 8601 Timestamp", "yyyy-mm-dd hh:mm:ss", "Full international date and 24-hour timestamp"),
            new DateFormatStandardItem("eu_short", "European Short Date", "dd/mm/yyyy", "Day / Month / Year format"),
            new DateFormatStandardItem("us_short", "US Short Date", "mm/dd/yyyy", "Month / Day / Year format"),
            new DateFormatStandardItem("std_long", "Standard Long Date", "dd mmmm yyyy", "Formal written date with full month name"),
            new DateFormatStandardItem("us_long", "US Long Date", "mmmm dd, yyyy", "Month Day, Year format"),
            new DateFormatStandardItem("mil_compact", "Financial / Military", "dd-mmm-yyyy", "Abbreviated month name (e.g. 28-Oct-2026)"),
            new DateFormatStandardItem("year_month", "Year and Month", "yyyy-mm", "Period indicator without day (YYYY-MM)"),
            new DateFormatStandardItem("full_weekday", "Weekday and Date", "dddd, mmmm dd, yyyy", "Full day of week and date")
        };

        public static IReadOnlyList<DateFormatStandardItem> GetStandards() => _standards;

        /// <summary>
        /// Attempts to parse any cell content (DateTime, Excel numeric serial, or text string) into a valid DateTime.
        /// </summary>
        public static bool TryParseDate(object? raw, out DateTime parsedDate)
        {
            parsedDate = DateTime.MinValue;
            if (raw == null) return false;

            if (raw is DateTime dt)
            {
                parsedDate = dt;
                return true;
            }

            if (raw is double d)
            {
                // Valid Excel date serials generally fall between 1 (1900-01-01) and 2958465 (9999-12-31)
                if (d >= 1.0 && d <= 2958465.0)
                {
                    try
                    {
                        parsedDate = DateTime.FromOADate(d);
                        return true;
                    }
                    catch { }
                }
            }

            string text = raw.ToString()?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(text) || text.Length < 4) return false;

            // Attempt exact matching with known format list
            if (DateTime.TryParseExact(text, RecognizedParseFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var exactDt))
            {
                parsedDate = exactDt;
                return true;
            }

            // Fallback to invariant culture parse
            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var invDt))
            {
                parsedDate = invDt;
                return true;
            }

            // Fallback to current culture parse
            if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out var curDt))
            {
                parsedDate = curDt;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Processes a 2D range matrix, converting text dates to real Excel date serials if requested.
        /// </summary>
        public static DateStandardizeResult StandardizeRange(
            object?[,] rawValues,
            string targetFormat,
            bool convertTextDates = true)
        {
            if (rawValues == null)
            {
                return new DateStandardizeResult(new object?[0, 0], 0, 0, 0, 0, targetFormat);
            }

            int rStart = rawValues.GetLowerBound(0);
            int rEnd = rawValues.GetUpperBound(0);
            int cStart = rawValues.GetLowerBound(1);
            int cEnd = rawValues.GetUpperBound(1);

            int rows = rEnd - rStart + 1;
            int cols = cEnd - cStart + 1;
            var processed = new object?[rows, cols];

            int totalCells = 0;
            int dateCellsFound = 0;
            int textDatesConverted = 0;
            int unchanged = 0;

            for (int r = rStart; r <= rEnd; r++)
            {
                int destR = r - rStart;
                for (int c = cStart; c <= cEnd; c++)
                {
                    int destC = c - cStart;
                    totalCells++;

                    object? cell = rawValues[r, c];
                    if (cell == null)
                    {
                        processed[destR, destC] = null;
                        unchanged++;
                        continue;
                    }

                    if (cell is DateTime)
                    {
                        dateCellsFound++;
                        processed[destR, destC] = cell;
                    }
                    else if (cell is double d && d >= 1.0 && d <= 2958465.0)
                    {
                        // Could be an existing Excel serial date
                        dateCellsFound++;
                        processed[destR, destC] = cell;
                    }
                    else if (convertTextDates && TryParseDate(cell, out var parsedDt))
                    {
                        // Text date converted to Excel date serial (OA date)
                        dateCellsFound++;
                        textDatesConverted++;
                        processed[destR, destC] = parsedDt.ToOADate();
                    }
                    else
                    {
                        processed[destR, destC] = cell;
                        unchanged++;
                    }
                }
            }

            return new DateStandardizeResult(
                processed,
                totalCells,
                dateCellsFound,
                textDatesConverted,
                unchanged,
                targetFormat);
        }
    }
}
