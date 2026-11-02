using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Ranges
{
    /// <summary>
    /// Core matrix transformation engine providing deterministic conversions between
    /// columns, rows, and multi-dimensional tabular ranges.
    /// Supports both 0-based and 1-based (Excel COM) source arrays.
    /// </summary>
    public static class RangeTransformer
    {
        /// <summary>
        /// Transforms the given source matrix according to the specified options.
        /// </summary>
        public static object?[,] Transform(object?[,] source, TransformRangeOptions options)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var (isValid, errorMessage) = options.Validate();
            if (!isValid)
            {
                throw new ArgumentException(errorMessage ?? "Invalid transformation options.", nameof(options));
            }

            switch (options.Type)
            {
                case TransformType.SingleColumnToRange:
                case TransformType.SingleRowToRange:
                    return TransformListToRange(source, options);

                case TransformType.RangeToSingleColumn:
                    return TransformRangeToSingleColumn(source, options.Order, options.SkipBlankCells);

                case TransformType.RangeToSingleRow:
                    return TransformRangeToSingleRow(source, options.Order, options.SkipBlankCells);

                default:
                    throw new NotSupportedException($"Transformation type '{options.Type}' is not supported.");
            }
        }

        /// <summary>
        /// Converts a single column or single row (or flattened sequence) into a multi-column range.
        /// </summary>
        public static object?[,] TransformListToRange(object?[,] source, TransformRangeOptions options)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));

            // Extract values sequentially in row/col order
            // If SingleColumnToRange: traverse down column first
            // If SingleRowToRange: traverse across row first
            TransformOrder order = (options.Type == TransformType.SingleRowToRange)
                ? TransformOrder.ByRow
                : TransformOrder.ByColumn;

            if (options.RecordSplit == TransformRecordSplit.BlankCellDelimited)
            {
                var rawItems = ExtractRawList(source, order);
                var records = new List<List<object?>>();
                var currentRecord = new List<object?>();

                foreach (var item in rawItems)
                {
                    if (IsBlank(item))
                    {
                        if (currentRecord.Count > 0)
                        {
                            records.Add(currentRecord);
                            currentRecord = new List<object?>();
                        }
                    }
                    else
                    {
                        currentRecord.Add(item);
                    }
                }

                if (currentRecord.Count > 0)
                {
                    records.Add(currentRecord);
                }

                if (records.Count == 0)
                {
                    return new object?[0, 0];
                }

                int maxCols = 0;
                foreach (var rec in records)
                {
                    if (rec.Count > maxCols) maxCols = rec.Count;
                }
                if (maxCols == 0) maxCols = 1;

                var result = new object?[records.Count, maxCols];
                for (int r = 0; r < records.Count; r++)
                {
                    for (int c = 0; c < records[r].Count; c++)
                    {
                        result[r, c] = records[r][c];
                    }
                }

                return result;
            }
            else // FixedValue
            {
                int cols = Math.Max(1, options.ColumnsPerRecord);
                var items = ExtractList(source, order, skipBlanks: false);

                if (items.Count == 0)
                {
                    return new object?[0, 0];
                }

                int rows = (int)Math.Ceiling((double)items.Count / cols);
                var result = new object?[rows, cols];

                for (int i = 0; i < items.Count; i++)
                {
                    int r = i / cols;
                    int c = i % cols;
                    result[r, c] = items[i];
                }

                return result;
            }
        }

        /// <summary>
        /// Flattens a 2D range into a single vertical column (N rows × 1 col).
        /// </summary>
        public static object?[,] TransformRangeToSingleColumn(object?[,] source, TransformOrder order, bool skipBlanks)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            var items = ExtractList(source, order, skipBlanks);
            if (items.Count == 0)
            {
                return new object?[0, 0];
            }

            var result = new object?[items.Count, 1];
            for (int r = 0; r < items.Count; r++)
            {
                result[r, 0] = items[r];
            }

            return result;
        }

        /// <summary>
        /// Flattens a 2D range into a single horizontal row (1 row × N cols).
        /// </summary>
        public static object?[,] TransformRangeToSingleRow(object?[,] source, TransformOrder order, bool skipBlanks)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            var items = ExtractList(source, order, skipBlanks);
            if (items.Count == 0)
            {
                return new object?[0, 0];
            }

            var result = new object?[1, items.Count];
            for (int c = 0; c < items.Count; c++)
            {
                result[0, c] = items[c];
            }

            return result;
        }

        /// <summary>
        /// Extracts elements from 2D matrix into a 1D list in the specified order,
        /// optionally filtering out blank cells.
        /// </summary>
        public static List<object?> ExtractList(object?[,] source, TransformOrder order, bool skipBlanks)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            int rStart = source.GetLowerBound(0);
            int rEnd = source.GetUpperBound(0);
            int cStart = source.GetLowerBound(1);
            int cEnd = source.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            var list = new List<object?>(rowCount * colCount);

            if (order == TransformOrder.ByRow)
            {
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        var val = source[rStart + r, cStart + c];
                        if (!skipBlanks || !IsBlank(val))
                        {
                            list.Add(val);
                        }
                    }
                }
            }
            else // ByColumn
            {
                for (int c = 0; c < colCount; c++)
                {
                    for (int r = 0; r < rowCount; r++)
                    {
                        var val = source[rStart + r, cStart + c];
                        if (!skipBlanks || !IsBlank(val))
                        {
                            list.Add(val);
                        }
                    }
                }
            }

            return list;
        }

        private static List<object?> ExtractRawList(object?[,] source, TransformOrder order)
        {
            return ExtractList(source, order, skipBlanks: false);
        }

        public static bool IsBlank(object? val)
        {
            if (val == null) return true;
            if (val is string s && string.IsNullOrWhiteSpace(s)) return true;
            return false;
        }
    }
}
