using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Navigation
{
    /// <summary>
    /// Color preset definition for Reading View highlights.
    /// </summary>
    public class ReadingViewColorPreset
    {
        public string Name { get; }
        public string FillHex { get; }
        public string BorderHex { get; }

        public ReadingViewColorPreset(string name, string fillHex, string borderHex)
        {
            Name = name;
            FillHex = fillHex;
            BorderHex = borderHex;
        }
    }

    /// <summary>
    /// Calculated geometry bounds for the Reading View highlight shapes in Excel points.
    /// </summary>
    public class HighlightGeometry
    {
        public bool HasRow { get; set; }
        public double RowLeft { get; set; }
        public double RowTop { get; set; }
        public double RowWidth { get; set; }
        public double RowHeight { get; set; }

        public bool HasColumn { get; set; }
        public double ColLeft { get; set; }
        public double ColTop { get; set; }
        public double ColWidth { get; set; }
        public double ColHeight { get; set; }

        public bool HasCell { get; set; }
        public double CellLeft { get; set; }
        public double CellTop { get; set; }
        public double CellWidth { get; set; }
        public double CellHeight { get; set; }
    }

    /// <summary>
    /// Engine performing mathematical geometry calculation, preset palette definition,
    /// and settings serialization for Reading View / Reading Layout crosshair highlighting.
    /// </summary>
    public static class ReadingLayoutEngine
    {
        public const string ShapeNameRow = "__EPS_ReadingView_Row__";
        public const string ShapeNameCol = "__EPS_ReadingView_Col__";
        public const string ShapeNameCell = "__EPS_ReadingView_Cell__";

        public static readonly IReadOnlyList<ReadingViewColorPreset> ColorPresets = new List<ReadingViewColorPreset>
        {
            new ReadingViewColorPreset("Sky Blue", "#38BDF8", "#0284C7"),
            new ReadingViewColorPreset("Emerald Green", "#34D399", "#059669"),
            new ReadingViewColorPreset("Warm Amber", "#FBBF24", "#D97706"),
            new ReadingViewColorPreset("Rose Pink", "#FB7185", "#E11D48"),
            new ReadingViewColorPreset("Lavender Purple", "#A78BFA", "#7C3AED"),
            new ReadingViewColorPreset("Slate Gray", "#94A3B8", "#475569")
        };

        private static readonly string SettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelProductivitySuite");
        private static readonly string SettingsFile = Path.Combine(SettingsDir, "ReadingViewSettings.json");

        /// <summary>
        /// Calculates exact geometric coordinates for the highlight overlay shapes.
        /// </summary>
        public static HighlightGeometry CalculateGeometry(
            double cellLeft,
            double cellTop,
            double cellWidth,
            double cellHeight,
            double maxLeft,
            double maxTop,
            double maxWidth,
            double maxHeight,
            ReadingViewSettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            var geo = new HighlightGeometry();

            double effectiveLeft = settings.Scope == ReadingHighlightScope.UsedRangeOnly ? maxLeft : 0;
            double effectiveWidth = settings.Scope == ReadingHighlightScope.UsedRangeOnly ? maxWidth : Math.Max(maxWidth, 4000);

            double effectiveTop = settings.Scope == ReadingHighlightScope.UsedRangeOnly ? maxTop : 0;
            double effectiveHeight = settings.Scope == ReadingHighlightScope.UsedRangeOnly ? maxHeight : Math.Max(maxHeight, 40000);

            bool includeRow = settings.Shape == ReadingHighlightShape.RowAndColumn || settings.Shape == ReadingHighlightShape.RowOnly;
            bool includeCol = settings.Shape == ReadingHighlightShape.RowAndColumn || settings.Shape == ReadingHighlightShape.ColumnOnly;
            bool includeCell = settings.Shape == ReadingHighlightShape.CellOnly || settings.Shape == ReadingHighlightShape.RowAndColumn;

            if (includeRow)
            {
                geo.HasRow = true;
                geo.RowLeft = effectiveLeft;
                geo.RowTop = cellTop;
                geo.RowWidth = effectiveWidth;
                geo.RowHeight = cellHeight;
            }

            if (includeCol)
            {
                geo.HasColumn = true;
                geo.ColLeft = cellLeft;
                geo.ColTop = effectiveTop;
                geo.ColWidth = cellWidth;
                geo.ColHeight = effectiveHeight;
            }

            if (includeCell)
            {
                geo.HasCell = true;
                geo.CellLeft = cellLeft;
                geo.CellTop = cellTop;
                geo.CellWidth = cellWidth;
                geo.CellHeight = cellHeight;
            }

            return geo;
        }

        /// <summary>
        /// Loads saved Reading View settings from disk, or returns default if not found.
        /// </summary>
        public static ReadingViewSettings LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    string json = File.ReadAllText(SettingsFile, Encoding.UTF8);
                    return DeserializeSettings(json);
                }
            }
            catch { }

            return ReadingViewSettings.CreateDefault();
        }

        /// <summary>
        /// Saves Reading View settings to persistent AppData storage.
        /// </summary>
        public static void SaveSettings(ReadingViewSettings settings)
        {
            if (settings == null) return;
            try
            {
                if (!Directory.Exists(SettingsDir))
                {
                    Directory.CreateDirectory(SettingsDir);
                }

                string json = SerializeSettings(settings);
                File.WriteAllText(SettingsFile, json, Encoding.UTF8);
            }
            catch { }
        }

        public static string SerializeSettings(ReadingViewSettings s)
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine($"  \"IsEnabled\": {(s.IsEnabled ? "true" : "false")},");
            sb.AppendLine($"  \"Shape\": {(int)s.Shape},");
            sb.AppendLine($"  \"Scope\": {(int)s.Scope},");
            sb.AppendLine($"  \"ColorHex\": \"{s.ColorHex}\",");
            sb.AppendLine($"  \"Transparency\": {s.Transparency.ToString("F2", CultureInfo.InvariantCulture)},");
            sb.AppendLine($"  \"BorderEnabled\": {(s.BorderEnabled ? "true" : "false")},");
            sb.AppendLine($"  \"BorderColorHex\": \"{s.BorderColorHex}\"");
            sb.Append("}");
            return sb.ToString();
        }

        public static ReadingViewSettings DeserializeSettings(string json)
        {
            var s = new ReadingViewSettings();
            if (string.IsNullOrWhiteSpace(json)) return s;

            var matchBool = Regex.Match(json, "\"IsEnabled\"\\s*:\\s*(true|false)", RegexOptions.IgnoreCase);
            if (matchBool.Success) s.IsEnabled = bool.Parse(matchBool.Groups[1].Value);

            var matchShape = Regex.Match(json, "\"Shape\"\\s*:\\s*(\\d+)");
            if (matchShape.Success && int.TryParse(matchShape.Groups[1].Value, out int shapeVal))
            {
                s.Shape = (ReadingHighlightShape)shapeVal;
            }

            var matchScope = Regex.Match(json, "\"Scope\"\\s*:\\s*(\\d+)");
            if (matchScope.Success && int.TryParse(matchScope.Groups[1].Value, out int scopeVal))
            {
                s.Scope = (ReadingHighlightScope)scopeVal;
            }

            var matchColor = Regex.Match(json, "\"ColorHex\"\\s*:\\s*\"([^\"]+)\"");
            if (matchColor.Success) s.ColorHex = matchColor.Groups[1].Value;

            var matchTrans = Regex.Match(json, "\"Transparency\"\\s*:\\s*([0-9.]+)");
            if (matchTrans.Success && double.TryParse(matchTrans.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double transVal))
            {
                s.Transparency = transVal;
            }

            var matchBorder = Regex.Match(json, "\"BorderEnabled\"\\s*:\\s*(true|false)", RegexOptions.IgnoreCase);
            if (matchBorder.Success) s.BorderEnabled = bool.Parse(matchBorder.Groups[1].Value);

            var matchBorderColor = Regex.Match(json, "\"BorderColorHex\"\\s*:\\s*\"([^\"]+)\"");
            if (matchBorderColor.Success) s.BorderColorHex = matchBorderColor.Groups[1].Value;

            return s;
        }
    }
}
