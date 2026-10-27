using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    public enum RecordResultCategory
    {
        SourceOnly,
        TargetOnly,
        Matching,
        SourceDuplicates,
        TargetDuplicates
    }

    public class RecordMatchItem
    {
        public int RowIndex { get; set; } // 1-based relative row
        public string CompositeKey { get; set; } = "";
        public object?[] Values { get; set; } = Array.Empty<object?>();
        public RecordResultCategory Category { get; set; }
    }

    public class RecordCompareOptions
    {
        public bool MatchCase { get; set; } = false;
        public bool TrimWhitespace { get; set; } = true;
        public bool HasHeaders { get; set; } = true;
        public int[] SourceKeyColumns { get; set; } = new[] { 0 }; // 0-based relative
        public int[] TargetKeyColumns { get; set; } = new[] { 0 }; // 0-based relative
    }

    public class RecordComparisonResult
    {
        public string[] SourceHeaders { get; set; } = Array.Empty<string>();
        public string[] TargetHeaders { get; set; } = Array.Empty<string>();

        public List<RecordMatchItem> SourceOnly { get; } = new List<RecordMatchItem>();
        public List<RecordMatchItem> TargetOnly { get; } = new List<RecordMatchItem>();
        public List<RecordMatchItem> Matching { get; } = new List<RecordMatchItem>();
        public List<RecordMatchItem> SourceDuplicates { get; } = new List<RecordMatchItem>();
        public List<RecordMatchItem> TargetDuplicates { get; } = new List<RecordMatchItem>();

        public int TotalSourceRows { get; set; }
        public int TotalTargetRows { get; set; }

        public object?[,] ExportToMatrix(RecordResultCategory category)
        {
            List<RecordMatchItem> items = category switch
            {
                RecordResultCategory.SourceOnly => SourceOnly,
                RecordResultCategory.TargetOnly => TargetOnly,
                RecordResultCategory.Matching => Matching,
                RecordResultCategory.SourceDuplicates => SourceDuplicates,
                RecordResultCategory.TargetDuplicates => TargetDuplicates,
                _ => SourceOnly
            };

            string[] headers = (category == RecordResultCategory.TargetOnly || category == RecordResultCategory.TargetDuplicates)
                ? TargetHeaders
                : SourceHeaders;

            int cols = headers.Length > 0 ? headers.Length : (items.Count > 0 ? items[0].Values.Length : 1);
            int rows = (headers.Length > 0 ? 1 : 0) + items.Count;

            object?[,] matrix = new object?[Math.Max(1, rows), cols];

            int currentRow = 0;
            if (headers.Length > 0)
            {
                for (int c = 0; c < headers.Length; c++)
                {
                    matrix[currentRow, c] = headers[c];
                }
                currentRow++;
            }

            for (int i = 0; i < items.Count; i++)
            {
                var valArray = items[i].Values;
                for (int c = 0; c < Math.Min(cols, valArray.Length); c++)
                {
                    matrix[currentRow + i, c] = valArray[c];
                }
            }

            return matrix;
        }
    }

    public static class RangeRecordComparer
    {
        public static RecordComparisonResult CompareRecords(object?[,] source, object?[,] target, RecordCompareOptions options)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var result = new RecordComparisonResult();

            int sRStart = source.GetLowerBound(0);
            int sREnd = source.GetUpperBound(0);
            int sCStart = source.GetLowerBound(1);
            int sCEnd = source.GetUpperBound(1);

            int tRStart = target.GetLowerBound(0);
            int tREnd = target.GetUpperBound(0);
            int tCStart = target.GetLowerBound(1);
            int tCEnd = target.GetUpperBound(1);

            int sCols = sCEnd - sCStart + 1;
            int tCols = tCEnd - tCStart + 1;

            int sDataStart = sRStart;
            int tDataStart = tRStart;

            if (options.HasHeaders)
            {
                result.SourceHeaders = new string[sCols];
                for (int c = 0; c < sCols; c++)
                {
                    result.SourceHeaders[c] = source[sRStart, sCStart + c]?.ToString() ?? $"Col_{c + 1}";
                }
                sDataStart++;

                result.TargetHeaders = new string[tCols];
                for (int c = 0; c < tCols; c++)
                {
                    result.TargetHeaders[c] = target[tRStart, tCStart + c]?.ToString() ?? $"Col_{c + 1}";
                }
                tDataStart++;
            }
            else
            {
                result.SourceHeaders = Enumerable.Range(1, sCols).Select(c => $"Col_{c}").ToArray();
                result.TargetHeaders = Enumerable.Range(1, tCols).Select(c => $"Col_{c}").ToArray();
            }

            var comparer = options.MatchCase ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;

            // Build Target Key Dictionary with frequency counts
            var targetKeyMap = new Dictionary<string, List<RecordMatchItem>>(comparer);
            int targetRowCount = 0;

            for (int r = tDataStart; r <= tREnd; r++)
            {
                targetRowCount++;
                string key = BuildKey(target, r, tCStart, tCEnd, options.TargetKeyColumns, options);
                var rowValues = GetRowValues(target, r, tCStart, tCEnd);
                var item = new RecordMatchItem
                {
                    RowIndex = r - tDataStart + 1,
                    CompositeKey = key,
                    Values = rowValues
                };

                if (!targetKeyMap.TryGetValue(key, out var list))
                {
                    list = new List<RecordMatchItem>();
                    targetKeyMap[key] = list;
                }
                else
                {
                    item.Category = RecordResultCategory.TargetDuplicates;
                    result.TargetDuplicates.Add(item);
                }
                list.Add(item);
            }
            result.TotalTargetRows = targetRowCount;

            // Process Source Rows
            var sourceKeyMap = new Dictionary<string, List<RecordMatchItem>>(comparer);
            int sourceRowCount = 0;

            for (int r = sDataStart; r <= sREnd; r++)
            {
                sourceRowCount++;
                string key = BuildKey(source, r, sCStart, sCEnd, options.SourceKeyColumns, options);
                var rowValues = GetRowValues(source, r, sCStart, sCEnd);
                var item = new RecordMatchItem
                {
                    RowIndex = r - sDataStart + 1,
                    CompositeKey = key,
                    Values = rowValues
                };

                if (!sourceKeyMap.TryGetValue(key, out var list))
                {
                    list = new List<RecordMatchItem>();
                    sourceKeyMap[key] = list;
                }
                else
                {
                    item.Category = RecordResultCategory.SourceDuplicates;
                    result.SourceDuplicates.Add(item);
                }
                list.Add(item);

                // Check against Target Key Map
                if (targetKeyMap.ContainsKey(key))
                {
                    item.Category = RecordResultCategory.Matching;
                    result.Matching.Add(item);
                }
                else
                {
                    item.Category = RecordResultCategory.SourceOnly;
                    result.SourceOnly.Add(item);
                }
            }
            result.TotalSourceRows = sourceRowCount;

            // Identify Target-Only rows (keys in target not in source)
            foreach (var kvp in targetKeyMap)
            {
                if (!sourceKeyMap.ContainsKey(kvp.Key))
                {
                    foreach (var targetItem in kvp.Value)
                    {
                        targetItem.Category = RecordResultCategory.TargetOnly;
                        result.TargetOnly.Add(targetItem);
                    }
                }
            }

            return result;
        }

        private static string BuildKey(object?[,] matrix, int row, int colStart, int colEnd, int[] keyColumns, RecordCompareOptions options)
        {
            var parts = new List<string>();
            if (keyColumns != null && keyColumns.Length > 0)
            {
                foreach (int colOffset in keyColumns)
                {
                    int c = colStart + colOffset;
                    if (c <= colEnd)
                    {
                        parts.Add(FormatCell(matrix[row, c], options));
                    }
                }
            }
            else
            {
                for (int c = colStart; c <= colEnd; c++)
                {
                    parts.Add(FormatCell(matrix[row, c], options));
                }
            }
            return string.Join("│", parts);
        }

        private static string FormatCell(object? val, RecordCompareOptions options)
        {
            if (val == null) return string.Empty;
            string str = val.ToString() ?? string.Empty;
            return options.TrimWhitespace ? str.Trim() : str;
        }

        private static object?[] GetRowValues(object?[,] matrix, int row, int colStart, int colEnd)
        {
            int cols = colEnd - colStart + 1;
            object?[] values = new object?[cols];
            for (int c = 0; c < cols; c++)
            {
                values[c] = matrix[row, colStart + c];
            }
            return values;
        }
    }
}
