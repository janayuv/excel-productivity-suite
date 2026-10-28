using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public enum RandomSampleMode
    {
        FixedCount,
        Percentage
    }

    public class RandomSampleOptions
    {
        public RandomSampleMode Mode { get; set; } = RandomSampleMode.FixedCount;
        public int Count { get; set; } = 10;
        public double Percentage { get; set; } = 10.0;
        public bool HasHeaders { get; set; } = true;
        public bool WithReplacement { get; set; } = false;
        public int? RandomSeed { get; set; } = null;
    }

    public class RandomSampleResult
    {
        public object?[,] SampleMatrix { get; }
        public IReadOnlyList<int> SelectedRowIndices { get; }
        public int TotalDataRows { get; }
        public int SampledRowsCount { get; }

        public RandomSampleResult(object?[,] sampleMatrix, IReadOnlyList<int> selectedIndices, int totalDataRows, int sampledCount)
        {
            SampleMatrix = sampleMatrix;
            SelectedRowIndices = selectedIndices;
            TotalDataRows = totalDataRows;
            SampledRowsCount = sampledCount;
        }
    }

    public static class RandomSampler
    {
        /// <summary>
        /// Extracts an unbiased random sample of rows from a dataset.
        /// </summary>
        public static RandomSampleResult ExtractSample(object?[,] matrix, RandomSampleOptions? options = null)
        {
            options ??= new RandomSampleOptions();

            if (matrix == null)
            {
                return new RandomSampleResult(new object?[0, 0], Array.Empty<int>(), 0, 0);
            }

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int totalRows = rEnd - rStart + 1;
            int totalCols = cEnd - cStart + 1;

            if (totalRows == 0 || totalCols == 0)
            {
                return new RandomSampleResult(new object?[0, 0], Array.Empty<int>(), 0, 0);
            }

            int dataStartRow = options.HasHeaders ? rStart + 1 : rStart;
            int totalDataRows = rEnd - dataStartRow + 1;

            if (totalDataRows <= 0)
            {
                return new RandomSampleResult(new object?[0, 0], Array.Empty<int>(), 0, 0);
            }

            // Determine sample size
            int sampleSize;
            if (options.Mode == RandomSampleMode.Percentage)
            {
                double pct = Math.Max(0.01, Math.Min(100.0, options.Percentage));
                sampleSize = (int)Math.Round((pct / 100.0) * totalDataRows);
                if (sampleSize < 1) sampleSize = 1;
            }
            else
            {
                sampleSize = Math.Max(1, options.Count);
            }

            if (!options.WithReplacement && sampleSize > totalDataRows)
            {
                sampleSize = totalDataRows;
            }

            var rand = options.RandomSeed.HasValue
                ? new Random(options.RandomSeed.Value)
                : new Random();

            var selectedDataIndices = new List<int>();

            if (options.WithReplacement)
            {
                for (int i = 0; i < sampleSize; i++)
                {
                    int chosen = rand.Next(dataStartRow, rEnd + 1);
                    selectedDataIndices.Add(chosen);
                }
            }
            else
            {
                // Fisher-Yates partial shuffle to select sampleSize distinct rows
                var pool = Enumerable.Range(dataStartRow, totalDataRows).ToArray();
                for (int i = 0; i < sampleSize; i++)
                {
                    int j = rand.Next(i, pool.Length);
                    int temp = pool[i];
                    pool[i] = pool[j];
                    pool[j] = temp;
                    selectedDataIndices.Add(pool[i]);
                }
            }

            // Construct output matrix
            int outRows = options.HasHeaders ? sampleSize + 1 : sampleSize;
            var outMatrix = new object?[outRows, totalCols];

            int outRowIdx = 0;
            if (options.HasHeaders)
            {
                // Copy headers
                for (int c = cStart; c <= cEnd; c++)
                {
                    outMatrix[0, c - cStart] = matrix[rStart, c];
                }
                outRowIdx = 1;
            }

            foreach (int srcR in selectedDataIndices)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    outMatrix[outRowIdx, c - cStart] = matrix[srcR, c];
                }
                outRowIdx++;
            }

            return new RandomSampleResult(outMatrix, selectedDataIndices, totalDataRows, sampleSize);
        }
    }
}
