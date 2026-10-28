using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelProductivitySuite.Engine.Analysis
{
    public class MetricItem
    {
        public string Name { get; set; }
        public string FormattedValue { get; set; }
        public string Category { get; set; }

        public MetricItem(string name, string formattedValue, string category)
        {
            Name = name;
            FormattedValue = formattedValue;
            Category = category;
        }
    }

    public class StatisticsResult
    {
        public int TotalCells { get; set; }
        public int NumericCells { get; set; }
        public int BlankCells { get; set; }
        public int TextCells { get; set; }
        public int ErrorCells { get; set; }
        public int DistinctNumericCount { get; set; }

        public double Sum { get; set; }
        public double Mean { get; set; }
        public double Median { get; set; }
        public double? Mode { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double Range { get; set; }

        public double SampleVariance { get; set; }
        public double PopulationVariance { get; set; }
        public double SampleStdDev { get; set; }
        public double PopulationStdDev { get; set; }
        public double? Skewness { get; set; }

        public List<MetricItem> MetricItems { get; } = new List<MetricItem>();
    }

    public static class QuickStatisticsEngine
    {
        public static StatisticsResult Calculate(object?[,]? data)
        {
            var result = new StatisticsResult();
            if (data == null)
            {
                PopulateMetricItems(result);
                return result;
            }

            int rStart = data.GetLowerBound(0);
            int rEnd = data.GetUpperBound(0);
            int cStart = data.GetLowerBound(1);
            int cEnd = data.GetUpperBound(1);

            int total = (rEnd - rStart + 1) * (cEnd - cStart + 1);
            result.TotalCells = total;

            var numbers = new List<double>(total);

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    object? val = data[r, c];
                    if (val == null || val is DBNull)
                    {
                        result.BlankCells++;
                    }
                    else if (val is string str)
                    {
                        if (string.IsNullOrWhiteSpace(str))
                        {
                            result.BlankCells++;
                        }
                        else if (str.StartsWith("#"))
                        {
                            result.ErrorCells++;
                        }
                        else if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedInv))
                        {
                            numbers.Add(parsedInv);
                        }
                        else if (double.TryParse(str, NumberStyles.Any, CultureInfo.CurrentCulture, out double parsedCurr))
                        {
                            numbers.Add(parsedCurr);
                        }
                        else
                        {
                            result.TextCells++;
                        }
                    }
                    else if (val is double d)
                    {
                        numbers.Add(d);
                    }
                    else if (val is int i)
                    {
                        numbers.Add(i);
                    }
                    else if (val is float f)
                    {
                        numbers.Add(f);
                    }
                    else if (val is decimal dec)
                    {
                        numbers.Add((double)dec);
                    }
                    else if (val is long l)
                    {
                        numbers.Add(l);
                    }
                    else if (val is short s)
                    {
                        numbers.Add(s);
                    }
                    else if (val is byte b)
                    {
                        numbers.Add(b);
                    }
                    else if (val is bool)
                    {
                        result.TextCells++;
                    }
                    else if (val is DateTime)
                    {
                        result.TextCells++;
                    }
                    else
                    {
                        string sVal = val.ToString() ?? "";
                        if (sVal.StartsWith("#"))
                        {
                            result.ErrorCells++;
                        }
                        else if (double.TryParse(sVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double p))
                        {
                            numbers.Add(p);
                        }
                        else
                        {
                            result.TextCells++;
                        }
                    }
                }
            }

            result.NumericCells = numbers.Count;

            if (numbers.Count > 0)
            {
                numbers.Sort();

                double sum = 0;
                for (int i = 0; i < numbers.Count; i++)
                {
                    sum += numbers[i];
                }

                result.Sum = sum;
                double mean = sum / numbers.Count;
                result.Mean = mean;

                // Median
                int n = numbers.Count;
                if (n % 2 == 1)
                {
                    result.Median = numbers[n / 2];
                }
                else
                {
                    result.Median = (numbers[(n / 2) - 1] + numbers[n / 2]) / 2.0;
                }

                result.Min = numbers[0];
                result.Max = numbers[n - 1];
                result.Range = result.Max - result.Min;

                // Mode
                var freq = new Dictionary<double, int>();
                int maxFreq = 0;
                foreach (double d in numbers)
                {
                    if (freq.TryGetValue(d, out int count))
                    {
                        freq[d] = count + 1;
                    }
                    else
                    {
                        freq[d] = 1;
                    }
                    if (freq[d] > maxFreq)
                    {
                        maxFreq = freq[d];
                    }
                }

                if (maxFreq > 1)
                {
                    result.Mode = freq.First(kv => kv.Value == maxFreq).Key;
                }
                else
                {
                    result.Mode = null;
                }

                result.DistinctNumericCount = freq.Count;

                // Variance & StdDev
                double sumSquaredDiffs = 0;
                for (int i = 0; i < n; i++)
                {
                    double diff = numbers[i] - mean;
                    sumSquaredDiffs += diff * diff;
                }

                result.PopulationVariance = sumSquaredDiffs / n;
                result.PopulationStdDev = Math.Sqrt(result.PopulationVariance);

                if (n > 1)
                {
                    result.SampleVariance = sumSquaredDiffs / (n - 1);
                    result.SampleStdDev = Math.Sqrt(result.SampleVariance);
                }
                else
                {
                    result.SampleVariance = 0;
                    result.SampleStdDev = 0;
                }

                // Skewness
                if (n >= 3 && result.SampleStdDev > 0)
                {
                    double sumCubedZ = 0;
                    for (int i = 0; i < n; i++)
                    {
                        double z = (numbers[i] - mean) / result.SampleStdDev;
                        sumCubedZ += z * z * z;
                    }
                    double factor = (double)n / ((n - 1) * (n - 2));
                    result.Skewness = factor * sumCubedZ;
                }
                else
                {
                    result.Skewness = null;
                }
            }

            PopulateMetricItems(result);
            return result;
        }

        private static void PopulateMetricItems(StatisticsResult res)
        {
            res.MetricItems.Clear();

            // Counts
            res.MetricItems.Add(new MetricItem("Total Cell Count", res.TotalCells.ToString("N0"), "Counts"));
            res.MetricItems.Add(new MetricItem("Numeric Cells", res.NumericCells.ToString("N0"), "Counts"));
            res.MetricItems.Add(new MetricItem("Blank Cells", res.BlankCells.ToString("N0"), "Counts"));
            res.MetricItems.Add(new MetricItem("Text Cells", res.TextCells.ToString("N0"), "Counts"));
            res.MetricItems.Add(new MetricItem("Error Cells", res.ErrorCells.ToString("N0"), "Counts"));
            res.MetricItems.Add(new MetricItem("Distinct Numbers", res.DistinctNumericCount.ToString("N0"), "Counts"));

            // Central Tendency
            if (res.NumericCells > 0)
            {
                res.MetricItems.Add(new MetricItem("Sum", FormatNumber(res.Sum), "Summary"));
                res.MetricItems.Add(new MetricItem("Average (Mean)", FormatNumber(res.Mean), "Central Tendency"));
                res.MetricItems.Add(new MetricItem("Median", FormatNumber(res.Median), "Central Tendency"));
                res.MetricItems.Add(new MetricItem("Mode", res.Mode.HasValue ? FormatNumber(res.Mode.Value) : "N/A", "Central Tendency"));

                // Extrema & Spread
                res.MetricItems.Add(new MetricItem("Minimum", FormatNumber(res.Min), "Dispersion"));
                res.MetricItems.Add(new MetricItem("Maximum", FormatNumber(res.Max), "Dispersion"));
                res.MetricItems.Add(new MetricItem("Range (Max - Min)", FormatNumber(res.Range), "Dispersion"));

                res.MetricItems.Add(new MetricItem("Sample Std. Deviation", FormatNumber(res.SampleStdDev), "Dispersion"));
                res.MetricItems.Add(new MetricItem("Population Std. Deviation", FormatNumber(res.PopulationStdDev), "Dispersion"));
                res.MetricItems.Add(new MetricItem("Sample Variance", FormatNumber(res.SampleVariance), "Dispersion"));
                res.MetricItems.Add(new MetricItem("Population Variance", FormatNumber(res.PopulationVariance), "Dispersion"));
                res.MetricItems.Add(new MetricItem("Skewness", res.Skewness.HasValue ? res.Skewness.Value.ToString("F4") : "N/A", "Dispersion"));
            }
        }

        private static string FormatNumber(double val)
        {
            if (Math.Abs(val) >= 1_000_000 || (Math.Abs(val) > 0 && Math.Abs(val) < 0.0001))
            {
                return val.ToString("G6");
            }
            if (Math.Abs(val % 1) < double.Epsilon)
            {
                return val.ToString("N0");
            }
            return val.ToString("N4").TrimEnd('0').TrimEnd('.');
        }

        public static object[,] GenerateSummaryTable(StatisticsResult stats, string? sourceAddress = null)
        {
            var rows = new List<(string Metric, object Value)>
            {
                ("Source Selection", sourceAddress ?? "Selection"),
                ("Total Cells", stats.TotalCells),
                ("Numeric Cells", stats.NumericCells),
                ("Blank Cells", stats.BlankCells),
                ("Text Cells", stats.TextCells),
                ("Error Cells", stats.ErrorCells),
                ("Distinct Numbers", stats.DistinctNumericCount)
            };

            if (stats.NumericCells > 0)
            {
                rows.Add(("Sum", stats.Sum));
                rows.Add(("Average (Mean)", stats.Mean));
                rows.Add(("Median", stats.Median));
                rows.Add(("Mode", stats.Mode.HasValue ? (object)stats.Mode.Value : "N/A"));
                rows.Add(("Minimum", stats.Min));
                rows.Add(("Maximum", stats.Max));
                rows.Add(("Range", stats.Range));
                rows.Add(("Sample Std. Dev.", stats.SampleStdDev));
                rows.Add(("Population Std. Dev.", stats.PopulationStdDev));
                rows.Add(("Sample Variance", stats.SampleVariance));
                rows.Add(("Population Variance", stats.PopulationVariance));
                rows.Add(("Skewness", stats.Skewness.HasValue ? (object)stats.Skewness.Value : "N/A"));
            }

            var table = new object[rows.Count + 1, 2];
            table[0, 0] = "Statistic";
            table[0, 1] = "Value";

            for (int i = 0; i < rows.Count; i++)
            {
                table[i + 1, 0] = rows[i].Metric;
                table[i + 1, 1] = rows[i].Value;
            }

            return table;
        }
    }
}
