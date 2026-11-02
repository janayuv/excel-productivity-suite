using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    public enum FrequencySortOrder
    {
        FrequencyDescending,
        FrequencyAscending,
        ValueAscending,
        ValueDescending
    }

    public class FrequencyDistributionOptions
    {
        public bool CaseSensitive { get; set; } = false;
        public bool IncludeBlanks { get; set; } = false;
        public string BlankLabel { get; set; } = "(Blank)";
        public FrequencySortOrder SortOrder { get; set; } = FrequencySortOrder.FrequencyDescending;
    }

    public class FrequencyRecord
    {
        public string Value { get; set; } = string.Empty;
        public int Count { get; set; }
        public double Percent { get; set; }
        public int CumulativeCount { get; set; }
        public double CumulativePercent { get; set; }

        public FrequencyRecord() { }

        public FrequencyRecord(string value, int count, double percent, int cumulativeCount, double cumulativePercent)
        {
            Value = value;
            Count = count;
            Percent = percent;
            CumulativeCount = cumulativeCount;
            CumulativePercent = cumulativePercent;
        }
    }

    public static class FrequencyDistributionCalculator
    {
        public static List<FrequencyRecord> Calculate(object?[,] matrix, FrequencyDistributionOptions? options = null)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            options ??= new FrequencyDistributionOptions();

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            var comparer = options.CaseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            var counts = new Dictionary<string, int>(comparer);
            int totalValid = 0;

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    object? cellVal = matrix[r, c];
                    if (cellVal == null || (cellVal is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        if (options.IncludeBlanks)
                        {
                            counts[options.BlankLabel] = counts.TryGetValue(options.BlankLabel, out int bCount) ? bCount + 1 : 1;
                            totalValid++;
                        }
                    }
                    else
                    {
                        string str = cellVal.ToString()!.Trim();
                        counts[str] = counts.TryGetValue(str, out int val) ? val + 1 : 1;
                        totalValid++;
                    }
                }
            }

            if (totalValid == 0) return new List<FrequencyRecord>();

            IEnumerable<KeyValuePair<string, int>> ordered = options.SortOrder switch
            {
                FrequencySortOrder.FrequencyDescending => counts.OrderByDescending(kvp => kvp.Value).ThenBy(kvp => kvp.Key, comparer),
                FrequencySortOrder.FrequencyAscending => counts.OrderBy(kvp => kvp.Value).ThenBy(kvp => kvp.Key, comparer),
                FrequencySortOrder.ValueAscending => counts.OrderBy(kvp => kvp.Key, comparer),
                FrequencySortOrder.ValueDescending => counts.OrderByDescending(kvp => kvp.Key, comparer),
                _ => counts.OrderByDescending(kvp => kvp.Value)
            };

            var records = new List<FrequencyRecord>(counts.Count);
            int runningCount = 0;

            foreach (var kvp in ordered)
            {
                runningCount += kvp.Value;
                double pct = Math.Round((double)kvp.Value / totalValid * 100.0, 2);
                double cumPct = Math.Round((double)runningCount / totalValid * 100.0, 2);

                records.Add(new FrequencyRecord(kvp.Key, kvp.Value, pct, runningCount, cumPct));
            }

            return records;
        }

        public static object?[,] GenerateMatrix(IReadOnlyList<FrequencyRecord> records, bool includeHeaders = true)
        {
            if (records == null || records.Count == 0)
            {
                if (!includeHeaders) return new object?[0, 0];
                var empty = new object?[1, 5];
                empty[0, 0] = "Value";
                empty[0, 1] = "Frequency";
                empty[0, 2] = "Percent (%)";
                empty[0, 3] = "Cumulative Count";
                empty[0, 4] = "Cumulative %";
                return empty;
            }

            int headerOffset = includeHeaders ? 1 : 0;
            int totalRows = records.Count + headerOffset;
            var matrix = new object?[totalRows, 5];

            if (includeHeaders)
            {
                matrix[0, 0] = "Value";
                matrix[0, 1] = "Frequency";
                matrix[0, 2] = "Percent (%)";
                matrix[0, 3] = "Cumulative Count";
                matrix[0, 4] = "Cumulative %";
            }

            for (int i = 0; i < records.Count; i++)
            {
                int r = i + headerOffset;
                var rec = records[i];
                matrix[r, 0] = rec.Value;
                matrix[r, 1] = rec.Count;
                matrix[r, 2] = $"{rec.Percent:F2}%";
                matrix[r, 3] = rec.CumulativeCount;
                matrix[r, 4] = $"{rec.CumulativePercent:F2}%";
            }

            return matrix;
        }
    }
}
