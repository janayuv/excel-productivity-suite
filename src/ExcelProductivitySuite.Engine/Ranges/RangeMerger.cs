using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum MergeMode
    {
        AppendRows,
        SideBySide
    }

    public class RangeMergeOptions
    {
        public MergeMode Mode { get; set; } = MergeMode.AppendRows;
        public bool HasHeaders { get; set; } = true;
        public bool AlignByHeaderNames { get; set; } = true;
        public bool IncludeHeadersOnce { get; set; } = true;
    }

    public class RangeMergeResult
    {
        public object?[,] MergedMatrix { get; set; } = new object?[0, 0];
        public int TotalRows => MergedMatrix.GetLength(0);
        public int TotalColumns => MergedMatrix.GetLength(1);
        public List<string> Warnings { get; } = new List<string>();
    }

    public static class RangeMerger
    {
        public static RangeMergeResult Merge(object?[,] range1, object?[,] range2, RangeMergeOptions options)
        {
            if (range1 == null) throw new ArgumentNullException(nameof(range1));
            if (range2 == null) throw new ArgumentNullException(nameof(range2));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var result = new RangeMergeResult();

            int r1RStart = range1.GetLowerBound(0);
            int r1REnd = range1.GetUpperBound(0);
            int r1CStart = range1.GetLowerBound(1);
            int r1CEnd = range1.GetUpperBound(1);

            int r2RStart = range2.GetLowerBound(0);
            int r2REnd = range2.GetUpperBound(0);
            int r2CStart = range2.GetLowerBound(1);
            int r2CEnd = range2.GetUpperBound(1);

            int r1Rows = r1REnd - r1RStart + 1;
            int r1Cols = r1CEnd - r1CStart + 1;
            int r2Rows = r2REnd - r2RStart + 1;
            int r2Cols = r2CEnd - r2CStart + 1;

            if (options.Mode == MergeMode.SideBySide)
            {
                int outRows = Math.Max(r1Rows, r2Rows);
                int outCols = r1Cols + r2Cols;
                object?[,] outMatrix = new object?[outRows, outCols];

                for (int r = 0; r < outRows; r++)
                {
                    // Copy from Range 1
                    for (int c = 0; c < r1Cols; c++)
                    {
                        outMatrix[r, c] = (r < r1Rows) ? range1[r1RStart + r, r1CStart + c] : null;
                    }

                    // Copy from Range 2
                    for (int c = 0; c < r2Cols; c++)
                    {
                        outMatrix[r, r1Cols + c] = (r < r2Rows) ? range2[r2RStart + r, r2CStart + c] : null;
                    }
                }

                if (r1Rows != r2Rows)
                {
                    result.Warnings.Add($"Row counts differ (Range 1: {r1Rows}, Range 2: {r2Rows}). The shorter range was padded with empty cells.");
                }

                result.MergedMatrix = outMatrix;
                return result;
            }

            // AppendRows Mode
            if (options.HasHeaders && options.AlignByHeaderNames)
            {
                // Extract headers
                var r1Headers = new List<string>();
                for (int c = 0; c < r1Cols; c++)
                {
                    r1Headers.Add(range1[r1RStart, r1CStart + c]?.ToString()?.Trim() ?? $"Column_{c + 1}");
                }

                var r2Headers = new List<string>();
                for (int c = 0; c < r2Cols; c++)
                {
                    r2Headers.Add(range2[r2RStart, r2CStart + c]?.ToString()?.Trim() ?? $"Column_{c + 1}");
                }

                // Unified ordered column list (Range 1 headers first, then any extra Range 2 headers)
                var unifiedHeaders = new List<string>(r1Headers);
                var headerSet = new HashSet<string>(r1Headers, StringComparer.OrdinalIgnoreCase);

                foreach (var h in r2Headers)
                {
                    if (headerSet.Add(h))
                    {
                        unifiedHeaders.Add(h);
                    }
                }

                int outCols = unifiedHeaders.Count;
                int r1DataRows = Math.Max(0, r1Rows - 1);
                int r2DataRows = Math.Max(0, r2Rows - 1);
                int outRows = 1 + r1DataRows + r2DataRows;

                object?[,] outMatrix = new object?[outRows, outCols];

                // Write header row
                for (int c = 0; c < outCols; c++)
                {
                    outMatrix[0, c] = unifiedHeaders[c];
                }

                // Map Range 1 columns to output
                for (int r = 0; r < r1DataRows; r++)
                {
                    for (int c = 0; c < r1Cols; c++)
                    {
                        string h = r1Headers[c];
                        int outCol = unifiedHeaders.FindIndex(x => string.Equals(x, h, StringComparison.OrdinalIgnoreCase));
                        if (outCol >= 0)
                        {
                            outMatrix[1 + r, outCol] = range1[r1RStart + 1 + r, r1CStart + c];
                        }
                    }
                }

                // Map Range 2 columns to output
                for (int r = 0; r < r2DataRows; r++)
                {
                    for (int c = 0; c < r2Cols; c++)
                    {
                        string h = r2Headers[c];
                        int outCol = unifiedHeaders.FindIndex(x => string.Equals(x, h, StringComparison.OrdinalIgnoreCase));
                        if (outCol >= 0)
                        {
                            outMatrix[1 + r1DataRows + r, outCol] = range2[r2RStart + 1 + r, r2CStart + c];
                        }
                    }
                }

                if (unifiedHeaders.Count > r1Cols)
                {
                    result.Warnings.Add($"Range 2 introduced {unifiedHeaders.Count - r1Cols} additional column(s) not in Range 1.");
                }

                result.MergedMatrix = outMatrix;
                return result;
            }
            else
            {
                // Positional stack
                int outCols = Math.Max(r1Cols, r2Cols);
                int r2StartRow = (options.HasHeaders && options.IncludeHeadersOnce) ? 1 : 0;
                int r2EffectiveRows = r2Rows - r2StartRow;

                int outRows = r1Rows + r2EffectiveRows;
                object?[,] outMatrix = new object?[outRows, outCols];

                // Copy Range 1
                for (int r = 0; r < r1Rows; r++)
                {
                    for (int c = 0; c < r1Cols; c++)
                    {
                        outMatrix[r, c] = range1[r1RStart + r, r1CStart + c];
                    }
                }

                // Copy Range 2
                for (int r = 0; r < r2EffectiveRows; r++)
                {
                    for (int c = 0; c < r2Cols; c++)
                    {
                        outMatrix[r1Rows + r, c] = range2[r2RStart + r2StartRow + r, r2CStart + c];
                    }
                }

                if (r1Cols != r2Cols)
                {
                    result.Warnings.Add($"Column counts differ (Range 1: {r1Cols}, Range 2: {r2Cols}). Output padded to {outCols} columns.");
                }

                result.MergedMatrix = outMatrix;
                return result;
            }
        }
    }
}
