using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum RangeCompareMode
    {
        CellByCell,
        FindValuesNotInTarget,
        FindCommonValues
    }

    public enum DifferenceType
    {
        ValueMismatch,
        MissingInSource,
        MissingInTarget,
        TypeMismatch
    }

    public class CellDifference
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string? SourceValue { get; set; }
        public string? TargetValue { get; set; }
        public DifferenceType DifferenceType { get; set; }
        public string Description { get; set; } = "";
    }

    public class RangeCompareOptions
    {
        public RangeCompareMode Mode { get; set; } = RangeCompareMode.CellByCell;
        public bool MatchCase { get; set; } = false;
        public bool TrimWhitespace { get; set; } = true;
        public bool IgnoreEmptyCells { get; set; } = false;
    }

    public class RangeComparisonResult
    {
        public int TotalCellsCompared { get; set; }
        public int TotalDifferences => Differences.Count;
        public List<CellDifference> Differences { get; } = new List<CellDifference>();
        public List<string> MissingValues { get; } = new List<string>();
        public List<string> CommonValues { get; } = new List<string>();
    }

    public static class RangeComparer
    {
        public static RangeComparisonResult Compare(object?[,] source, object?[,] target, RangeCompareOptions options)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var result = new RangeComparisonResult();

            int sRowStart = source.GetLowerBound(0);
            int sRowEnd = source.GetUpperBound(0);
            int sColStart = source.GetLowerBound(1);
            int sColEnd = source.GetUpperBound(1);

            int tRowStart = target.GetLowerBound(0);
            int tRowEnd = target.GetUpperBound(0);
            int tColStart = target.GetLowerBound(1);
            int tColEnd = target.GetUpperBound(1);

            int sRows = sRowEnd - sRowStart + 1;
            int sCols = sColEnd - sColStart + 1;
            int tRows = tRowEnd - tRowStart + 1;
            int tCols = tColEnd - tColStart + 1;

            var stringComparer = options.MatchCase ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;

            if (options.Mode == RangeCompareMode.CellByCell)
            {
                int maxRows = Math.Max(sRows, tRows);
                int maxCols = Math.Max(sCols, tCols);
                result.TotalCellsCompared = maxRows * maxCols;

                for (int r = 0; r < maxRows; r++)
                {
                    for (int c = 0; c < maxCols; c++)
                    {
                        bool hasSource = r < sRows && c < sCols;
                        bool hasTarget = r < tRows && c < tCols;

                        object? sVal = hasSource ? source[sRowStart + r, sColStart + c] : null;
                        object? tVal = hasTarget ? target[tRowStart + r, tColStart + c] : null;

                        string sStr = Normalize(sVal, options);
                        string tStr = Normalize(tVal, options);

                        if (options.IgnoreEmptyCells && string.IsNullOrEmpty(sStr) && string.IsNullOrEmpty(tStr))
                        {
                            continue;
                        }

                        if (!hasSource && hasTarget)
                        {
                            result.Differences.Add(new CellDifference
                            {
                                Row = r + 1,
                                Column = c + 1,
                                SourceValue = null,
                                TargetValue = tStr,
                                DifferenceType = DifferenceType.MissingInSource,
                                Description = $"Missing in source; target has '{tStr}'"
                            });
                        }
                        else if (hasSource && !hasTarget)
                        {
                            result.Differences.Add(new CellDifference
                            {
                                Row = r + 1,
                                Column = c + 1,
                                SourceValue = sStr,
                                TargetValue = null,
                                DifferenceType = DifferenceType.MissingInTarget,
                                Description = $"Source has '{sStr}'; missing in target"
                            });
                        }
                        else if (!stringComparer.Equals(sStr, tStr))
                        {
                            result.Differences.Add(new CellDifference
                            {
                                Row = r + 1,
                                Column = c + 1,
                                SourceValue = sStr,
                                TargetValue = tStr,
                                DifferenceType = DifferenceType.ValueMismatch,
                                Description = $"Mismatch: '{sStr}' vs '{tStr}'"
                            });
                        }
                    }
                }
            }
            else if (options.Mode == RangeCompareMode.FindValuesNotInTarget)
            {
                var targetSet = new HashSet<string>(stringComparer);
                for (int r = tRowStart; r <= tRowEnd; r++)
                {
                    for (int c = tColStart; c <= tColEnd; c++)
                    {
                        string val = Normalize(target[r, c], options);
                        if (!string.IsNullOrEmpty(val) || !options.IgnoreEmptyCells)
                        {
                            targetSet.Add(val);
                        }
                    }
                }

                int compared = 0;
                for (int r = sRowStart; r <= sRowEnd; r++)
                {
                    for (int c = sColStart; c <= sColEnd; c++)
                    {
                        compared++;
                        string sVal = Normalize(source[r, c], options);
                        if (options.IgnoreEmptyCells && string.IsNullOrEmpty(sVal)) continue;

                        if (!targetSet.Contains(sVal))
                        {
                            result.MissingValues.Add(sVal);
                            result.Differences.Add(new CellDifference
                            {
                                Row = r - sRowStart + 1,
                                Column = c - sColStart + 1,
                                SourceValue = sVal,
                                TargetValue = null,
                                DifferenceType = DifferenceType.MissingInTarget,
                                Description = $"Value '{sVal}' not found in target range"
                            });
                        }
                    }
                }
                result.TotalCellsCompared = compared;
            }
            else if (options.Mode == RangeCompareMode.FindCommonValues)
            {
                var targetSet = new HashSet<string>(stringComparer);
                for (int r = tRowStart; r <= tRowEnd; r++)
                {
                    for (int c = tColStart; c <= tColEnd; c++)
                    {
                        string val = Normalize(target[r, c], options);
                        if (!string.IsNullOrEmpty(val) || !options.IgnoreEmptyCells)
                        {
                            targetSet.Add(val);
                        }
                    }
                }

                int compared = 0;
                var commonFound = new HashSet<string>(stringComparer);

                for (int r = sRowStart; r <= sRowEnd; r++)
                {
                    for (int c = sColStart; c <= sColEnd; c++)
                    {
                        compared++;
                        string sVal = Normalize(source[r, c], options);
                        if (options.IgnoreEmptyCells && string.IsNullOrEmpty(sVal)) continue;

                        if (targetSet.Contains(sVal) && commonFound.Add(sVal))
                        {
                            result.CommonValues.Add(sVal);
                        }
                    }
                }
                result.TotalCellsCompared = compared;
            }

            return result;
        }

        private static string Normalize(object? val, RangeCompareOptions options)
        {
            if (val == null) return string.Empty;
            string str = val.ToString() ?? string.Empty;
            return options.TrimWhitespace ? str.Trim() : str;
        }
    }
}
