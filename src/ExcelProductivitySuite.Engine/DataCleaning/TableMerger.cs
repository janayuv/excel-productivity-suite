using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Engine.DataCleaning
{
    public class TableMergeOptions
    {
        public int PrimaryKeyColumn { get; set; } = 0;
        public int LookupKeyColumn { get; set; } = 0;
        public bool HasHeaders { get; set; } = true;
        public bool AppendUnmatchedRows { get; set; } = false;
        public bool MatchCase { get; set; } = false;
    }

    /// <summary>
    /// Pure computational engine for joining and updating tables based on key-column matching.
    /// </summary>
    public static class TableMerger
    {
        public static object?[,] Merge(
            object?[,] primaryTable,
            object?[,] lookupTable,
            TableMergeOptions options,
            out int matchedCount,
            out int unmatchedCount)
        {
            if (primaryTable == null) throw new ArgumentNullException(nameof(primaryTable));
            if (lookupTable == null) throw new ArgumentNullException(nameof(lookupTable));
            if (options == null) throw new ArgumentNullException(nameof(options));

            matchedCount = 0;
            unmatchedCount = 0;

            int pRowStart = primaryTable.GetLowerBound(0);
            int pRowEnd = primaryTable.GetUpperBound(0);
            int pColStart = primaryTable.GetLowerBound(1);
            int pColEnd = primaryTable.GetUpperBound(1);

            int lRowStart = lookupTable.GetLowerBound(0);
            int lRowEnd = lookupTable.GetUpperBound(0);
            int lColStart = lookupTable.GetLowerBound(1);
            int lColEnd = lookupTable.GetUpperBound(1);

            int pRows = pRowEnd - pRowStart + 1;
            int pCols = pColEnd - pColStart + 1;
            int lRows = lRowEnd - lRowStart + 1;
            int lCols = lColEnd - lColStart + 1;

            if (pRows <= 0 || pCols <= 0) return (object?[,])lookupTable.Clone();
            if (lRows <= 0 || lCols <= 0) return (object?[,])primaryTable.Clone();

            int pDataStart = options.HasHeaders ? pRowStart + 1 : pRowStart;
            int lDataStart = options.HasHeaders ? lRowStart + 1 : lRowStart;

            int pKeyCol = pColStart + Math.Min(pCols - 1, Math.Max(0, options.PrimaryKeyColumn));
            int lKeyCol = lColStart + Math.Min(lCols - 1, Math.Max(0, options.LookupKeyColumn));

            // Which columns from lookup table are appended? All columns except the lookup key column.
            var lookupColsToAppend = new List<int>();
            for (int c = lColStart; c <= lColEnd; c++)
            {
                if (c != lKeyCol) lookupColsToAppend.Add(c);
            }

            int finalCols = pCols + lookupColsToAppend.Count;

            var comparer = options.MatchCase ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            // Map lookup key -> lookup row index
            var lookupMap = new Dictionary<string, int>(comparer);
            for (int r = lDataStart; r <= lRowEnd; r++)
            {
                string key = lookupTable[r, lKeyCol]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(key) && !lookupMap.ContainsKey(key))
                {
                    lookupMap[key] = r;
                }
            }

            // Track matched lookup rows
            var matchedLookupRows = new HashSet<int>();

            // Calculate total rows for result
            int totalResultRows = pRows;
            if (options.AppendUnmatchedRows)
            {
                int totalLookupDataRows = lRowEnd - lDataStart + 1;
                totalResultRows += totalLookupDataRows; // upper bound
            }

            object?[,] result = new object?[totalResultRows, finalCols];
            int destRow = 0;

            // Handle Header row
            if (options.HasHeaders)
            {
                for (int c = 0; c < pCols; c++)
                {
                    result[0, c] = primaryTable[pRowStart, pColStart + c];
                }

                for (int i = 0; i < lookupColsToAppend.Count; i++)
                {
                    result[0, pCols + i] = lookupTable[lRowStart, lookupColsToAppend[i]];
                }
                destRow = 1;
            }

            // Process Primary Table rows
            for (int r = pDataStart; r <= pRowEnd; r++)
            {
                // Copy primary table columns
                for (int c = 0; c < pCols; c++)
                {
                    result[destRow, c] = primaryTable[r, pColStart + c];
                }

                string key = primaryTable[r, pKeyCol]?.ToString()?.Trim() ?? "";
                if (!string.IsNullOrEmpty(key) && lookupMap.TryGetValue(key, out int lRow))
                {
                    matchedCount++;
                    matchedLookupRows.Add(lRow);

                    // Copy appended columns from lookup row
                    for (int i = 0; i < lookupColsToAppend.Count; i++)
                    {
                        result[destRow, pCols + i] = lookupTable[lRow, lookupColsToAppend[i]];
                    }
                }
                else
                {
                    unmatchedCount++;
                }

                destRow++;
            }

            // Handle Unmatched lookup rows if configured
            if (options.AppendUnmatchedRows)
            {
                for (int r = lDataStart; r <= lRowEnd; r++)
                {
                    if (!matchedLookupRows.Contains(r))
                    {
                        // Put lookup key in primary key column
                        int keyOffset = pKeyCol - pColStart;
                        result[destRow, keyOffset] = lookupTable[r, lKeyCol];

                        // Copy other columns
                        for (int i = 0; i < lookupColsToAppend.Count; i++)
                        {
                            result[destRow, pCols + i] = lookupTable[r, lookupColsToAppend[i]];
                        }
                        destRow++;
                    }
                }
            }

            // Trim if destRow < totalResultRows
            if (destRow < totalResultRows)
            {
                object?[,] trimmed = new object?[destRow, finalCols];
                for (int r = 0; r < destRow; r++)
                {
                    for (int c = 0; c < finalCols; c++)
                    {
                        trimmed[r, c] = result[r, c];
                    }
                }
                return trimmed;
            }

            return result;
        }
    }
}
