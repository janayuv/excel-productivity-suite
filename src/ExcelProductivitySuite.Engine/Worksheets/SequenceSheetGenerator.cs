using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelProductivitySuite.Engine.Worksheets
{
    public enum SequenceType
    {
        Numbers,
        Months,
        Days,
        Quarters
    }

    public class SequenceSheetOptions
    {
        public SequenceType Type { get; set; } = SequenceType.Numbers;
        public string Prefix { get; set; } = "Sheet_";
        public string Suffix { get; set; } = "";
        public int StartNumber { get; set; } = 1;
        public int Count { get; set; } = 5;
        public int Step { get; set; } = 1;
        public int ZeroPadding { get; set; } = 2;
        public bool FullName { get; set; } = false;
        public int Year { get; set; } = DateTime.Today.Year;
    }

    public static class SequenceSheetGenerator
    {
        private static readonly string[] ShortMonths = DateTimeFormatInfo.InvariantInfo.AbbreviatedMonthNames.Take(12).ToArray();
        private static readonly string[] FullMonths = DateTimeFormatInfo.InvariantInfo.MonthNames.Take(12).ToArray();

        private static readonly string[] ShortDays = new[] { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun" };
        private static readonly string[] FullDays = new[] { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" };

        public static List<string> GenerateNames(SequenceSheetOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            var names = new List<string>();
            int count = Math.Max(1, Math.Min(options.Count, 255));

            switch (options.Type)
            {
                case SequenceType.Numbers:
                    int cur = options.StartNumber;
                    int step = options.Step == 0 ? 1 : options.Step;
                    for (int i = 0; i < count; i++)
                    {
                        string numStr = options.ZeroPadding > 1 
                            ? cur.ToString($"D{options.ZeroPadding}") 
                            : cur.ToString();
                        names.Add($"{options.Prefix}{numStr}{options.Suffix}");
                        cur += step;
                    }
                    break;

                case SequenceType.Months:
                    var monthList = options.FullName ? FullMonths : ShortMonths;
                    int startMonth = Math.Max(1, Math.Min(options.StartNumber, 12)) - 1;
                    for (int i = 0; i < count; i++)
                    {
                        int mIdx = (startMonth + i) % 12;
                        names.Add($"{options.Prefix}{monthList[mIdx]}{options.Suffix}");
                    }
                    break;

                case SequenceType.Days:
                    var dayList = options.FullName ? FullDays : ShortDays;
                    int startDay = Math.Max(1, Math.Min(options.StartNumber, 7)) - 1;
                    for (int i = 0; i < count; i++)
                    {
                        int dIdx = (startDay + i) % 7;
                        names.Add($"{options.Prefix}{dayList[dIdx]}{options.Suffix}");
                    }
                    break;

                case SequenceType.Quarters:
                    int startQ = Math.Max(1, Math.Min(options.StartNumber, 4));
                    for (int i = 0; i < count; i++)
                    {
                        int qNum = ((startQ - 1 + i) % 4) + 1;
                        int yOffset = (startQ - 1 + i) / 4;
                        string qStr = $"Q{qNum}";
                        string yStr = options.Year > 0 ? $"_{options.Year + yOffset}" : "";
                        names.Add($"{options.Prefix}{qStr}{yStr}{options.Suffix}");
                    }
                    break;
            }

            return names;
        }

        public static (List<string> ValidNames, List<string> Errors) ValidateGeneratedNames(
            IEnumerable<string> proposedNames, 
            IEnumerable<string>? existingNames = null)
        {
            var valid = new List<string>();
            var errors = new List<string>();
            var existingSet = new HashSet<string>(existingNames ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var name in proposedNames)
            {
                var (isValid, err) = WorksheetNameValidator.ValidateSheetName(name, existingSet);
                if (!isValid)
                {
                    errors.Add($"'{name}': {err}");
                }
                else if (seen.Contains(name))
                {
                    errors.Add($"'{name}': Duplicate proposed name in sequence.");
                }
                else
                {
                    seen.Add(name);
                    existingSet.Add(name);
                    valid.Add(name);
                }
            }

            return (valid, errors);
        }
    }
}
