using System;
using System.Collections.Generic;
using System.Linq;
using ExcelProductivitySuite.Core.Abstractions;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    public class ColorStatisticGroup
    {
        public int ColorRgb { get; set; }
        public string HexColor => $"#{ColorRgb:X6}";
        public int CellCount { get; set; }
        public int NumericCount { get; set; }
        public double Sum { get; set; }
        public double Average => NumericCount > 0 ? Sum / NumericCount : 0.0;
        public double? Min { get; set; }
        public double? Max { get; set; }
    }

    /// <summary>
    /// Engine for filtering and calculating summary statistics grouped by cell background or font color.
    /// </summary>
    public static class ColorAnalyzer
    {
        public static List<(int Row, int Column)> FilterCellsByColor(
            IEnumerable<CellColorInfo> cells,
            int targetColorRgb,
            bool isFillColor)
        {
            if (cells == null) throw new ArgumentNullException(nameof(cells));

            var results = new List<(int, int)>();
            foreach (var c in cells)
            {
                int color = isFillColor ? c.FillColorRgb : c.FontColorRgb;
                if (color == targetColorRgb)
                {
                    results.Add((c.Row, c.Column));
                }
            }
            return results;
        }

        public static List<ColorStatisticGroup> CalculateColorStatistics(
            IEnumerable<CellColorInfo> cells,
            bool isFillColor)
        {
            if (cells == null) throw new ArgumentNullException(nameof(cells));

            var groups = new Dictionary<int, ColorStatisticGroup>();

            foreach (var c in cells)
            {
                int color = isFillColor ? c.FillColorRgb : c.FontColorRgb;

                if (!groups.TryGetValue(color, out var stat))
                {
                    stat = new ColorStatisticGroup { ColorRgb = color };
                    groups[color] = stat;
                }

                stat.CellCount++;

                if (TryGetNumeric(c.Value, out double d))
                {
                    stat.NumericCount++;
                    stat.Sum += d;
                    stat.Min = stat.Min.HasValue ? Math.Min(stat.Min.Value, d) : d;
                    stat.Max = stat.Max.HasValue ? Math.Max(stat.Max.Value, d) : d;
                }
            }

            return groups.Values.OrderByDescending(g => g.CellCount).ToList();
        }

        private static bool TryGetNumeric(object? val, out double d)
        {
            d = 0;
            if (val == null) return false;
            if (val is double dbl) { d = dbl; return true; }
            if (val is int i) { d = i; return true; }
            if (val is float f) { d = f; return true; }
            if (val is decimal dec) { d = (double)dec; return true; }
            return double.TryParse(val.ToString(), out d);
        }
    }
}
