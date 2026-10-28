using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelProductivitySuite.Engine.Analysis
{
    public class TargetVsActualConfig
    {
        public int CategoryColumnIndex { get; set; } = 1;
        public int ActualColumnIndex { get; set; } = 2;
        public int TargetColumnIndex { get; set; } = 3;
        public bool HasHeaderRow { get; set; } = true;
        public string ChartTitle { get; set; } = "Target vs. Actual Performance";

        public int ValidDataRowCount { get; set; }
        public double TotalActual { get; set; }
        public double TotalTarget { get; set; }
        public double TotalVariance => TotalActual - TotalTarget;
        public double VariancePercent => Math.Abs(TotalTarget) > double.Epsilon ? (TotalVariance / TotalTarget) * 100.0 : 0.0;
    }

    public static class TargetVsActualDetector
    {
        private static readonly string[] ActualKeywords = new[] { "actual", "sales", "achieved", "result", "real", "current", "revenue" };
        private static readonly string[] TargetKeywords = new[] { "target", "budget", "goal", "plan", "quota", "forecast", "benchmark" };
        private static readonly string[] CategoryKeywords = new[] { "category", "month", "quarter", "name", "item", "product", "region", "rep", "period" };

        public static TargetVsActualConfig Detect(object?[,]? data, int colCount)
        {
            var config = new TargetVsActualConfig();
            if (data == null || colCount < 2) return config;

            int rStart = data.GetLowerBound(0);
            int rEnd = data.GetUpperBound(0);
            int cStart = data.GetLowerBound(1);
            int cEnd = data.GetUpperBound(1);

            int actualCols = cEnd - cStart + 1;
            if (actualCols < 2) return config;

            // Check if first row has headers
            string[] headers = new string[actualCols];
            for (int c = 0; c < actualCols; c++)
            {
                headers[c] = data[rStart, cStart + c]?.ToString()?.Trim() ?? $"Column {c + 1}";
            }

            int catCol = -1;
            int actCol = -1;
            int tgtCol = -1;

            // Scan headers for matches
            for (int c = 0; c < actualCols; c++)
            {
                string hLower = headers[c].ToLowerInvariant();

                if (actCol == -1 && MatchesAny(hLower, ActualKeywords))
                {
                    actCol = c + 1;
                    continue;
                }

                if (tgtCol == -1 && MatchesAny(hLower, TargetKeywords))
                {
                    tgtCol = c + 1;
                    continue;
                }

                if (catCol == -1 && MatchesAny(hLower, CategoryKeywords))
                {
                    catCol = c + 1;
                    continue;
                }
            }

            // Fallback column positions
            if (catCol == -1) catCol = 1;

            if (actCol == -1)
            {
                actCol = (catCol == 1) ? 2 : 1;
            }

            if (tgtCol == -1)
            {
                for (int c = 1; c <= actualCols; c++)
                {
                    if (c != catCol && c != actCol)
                    {
                        tgtCol = c;
                        break;
                    }
                }
                if (tgtCol == -1) tgtCol = (actCol == 2) ? 3 : 2;
            }

            config.CategoryColumnIndex = catCol;
            config.ActualColumnIndex = actCol;
            config.TargetColumnIndex = tgtCol;

            // Compute summary
            int dataStartRow = rStart + 1; // assume row 1 is header
            int validRows = 0;
            double totAct = 0;
            double totTgt = 0;

            for (int r = dataStartRow; r <= rEnd; r++)
            {
                double? actVal = ParseDouble(data[r, cStart + actCol - 1]);
                double? tgtVal = ParseDouble(data[r, cStart + tgtCol - 1]);

                if (actVal.HasValue || tgtVal.HasValue)
                {
                    validRows++;
                    if (actVal.HasValue) totAct += actVal.Value;
                    if (tgtVal.HasValue) totTgt += tgtVal.Value;
                }
            }

            config.ValidDataRowCount = validRows;
            config.TotalActual = totAct;
            config.TotalTarget = totTgt;

            return config;
        }

        private static bool MatchesAny(string header, string[] keywords)
        {
            foreach (var kw in keywords)
            {
                if (header.Contains(kw)) return true;
            }
            return false;
        }

        private static double? ParseDouble(object? val)
        {
            if (val == null || val is DBNull) return null;
            if (val is double d) return d;
            if (val is int i) return i;
            if (val is float f) return f;
            if (val is decimal dec) return (double)dec;
            if (val is long l) return l;

            string str = val.ToString()?.Trim() ?? "";
            if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double p1)) return p1;
            if (double.TryParse(str, NumberStyles.Any, CultureInfo.CurrentCulture, out double p2)) return p2;

            return null;
        }
    }
}
