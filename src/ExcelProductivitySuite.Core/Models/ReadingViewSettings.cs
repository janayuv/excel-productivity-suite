using System;
using System.Globalization;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Geometric shape of the Reading View active cell highlight.
    /// </summary>
    public enum ReadingHighlightShape
    {
        RowAndColumn = 0,
        RowOnly = 1,
        ColumnOnly = 2,
        CellOnly = 3
    }

    /// <summary>
    /// Scope / boundary of the Reading View crosshair highlight.
    /// </summary>
    public enum ReadingHighlightScope
    {
        EntireRowColumn = 0,
        UsedRangeOnly = 1
    }

    /// <summary>
    /// Configurable user settings for the Reading View / Reading Layout crosshair highlighter.
    /// </summary>
    public class ReadingViewSettings
    {
        public bool IsEnabled { get; set; } = false;
        public ReadingHighlightShape Shape { get; set; } = ReadingHighlightShape.RowAndColumn;
        public ReadingHighlightScope Scope { get; set; } = ReadingHighlightScope.EntireRowColumn;

        public string ColorHex { get; set; } = "#38BDF8"; // Sky Blue
        public double Transparency { get; set; } = 0.75;  // 75% transparent (25% opacity)
        public bool BorderEnabled { get; set; } = true;
        public string BorderColorHex { get; set; } = "#0284C7"; // Darker Blue outline

        public int ColorRgb => HexToRgb(ColorHex);
        public int BorderColorRgb => HexToRgb(BorderColorHex);

        public ReadingViewSettings Clone()
        {
            return new ReadingViewSettings
            {
                IsEnabled = this.IsEnabled,
                Shape = this.Shape,
                Scope = this.Scope,
                ColorHex = this.ColorHex,
                Transparency = this.Transparency,
                BorderEnabled = this.BorderEnabled,
                BorderColorHex = this.BorderColorHex
            };
        }

        public static ReadingViewSettings CreateDefault()
        {
            return new ReadingViewSettings();
        }

        public static int HexToRgb(string? hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return 0x00E0F2FE;

            string clean = hex!.Trim().TrimStart('#');
            if (clean.Length == 6 && int.TryParse(clean, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int hexVal))
            {
                int r = (hexVal >> 16) & 0xFF;
                int g = (hexVal >> 8) & 0xFF;
                int b = hexVal & 0xFF;
                // Excel COM RGB format: b << 16 | g << 8 | r
                return (b << 16) | (g << 8) | r;
            }
            return 0x00E0F2FE;
        }

        public static string RgbToHex(int rgb)
        {
            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            return $"#{r:X2}{g:X2}{b:X2}";
        }
    }
}
