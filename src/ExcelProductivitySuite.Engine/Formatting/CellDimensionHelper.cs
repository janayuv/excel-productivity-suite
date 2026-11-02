using System;
using ExcelProductivitySuite.Core.Abstractions;

namespace ExcelProductivitySuite.Engine.Formatting
{
    public static class CellDimensionHelper
    {
        public const double PointsPerInch = 72.0;
        public const double PointsPerCm = 72.0 / 2.54; // ~28.3464567

        public static double ToPoints(double value, CellDimensionUnit unit)
        {
            return unit switch
            {
                CellDimensionUnit.Points => value,
                CellDimensionUnit.Inches => value * PointsPerInch,
                CellDimensionUnit.Centimeters => value * PointsPerCm,
                _ => value
            };
        }

        public static double FromPoints(double points, CellDimensionUnit unit)
        {
            return unit switch
            {
                CellDimensionUnit.Points => points,
                CellDimensionUnit.Inches => points / PointsPerInch,
                CellDimensionUnit.Centimeters => points / PointsPerCm,
                _ => points
            };
        }

        public static double ClampRowHeight(double points)
        {
            return Math.Min(409.5, Math.Max(0.0, points));
        }

        public static double ClampColumnWidth(double points)
        {
            // Maximum column width in Excel is 255 characters (~1340 points)
            return Math.Min(1340.0, Math.Max(0.0, points));
        }
    }
}
