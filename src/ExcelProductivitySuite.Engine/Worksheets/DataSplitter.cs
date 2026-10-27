using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelProductivitySuite.Engine.Worksheets
{
    public class DataSplitOptions
    {
        public int SplitColumnIndex { get; set; } = 0; // 0-based relative to selection
        public bool HasHeaders { get; set; } = true;
        public bool TrimKeyValues { get; set; } = true;
        public bool CaseSensitive { get; set; } = false;
        public string SheetNamePrefix { get; set; } = "";
    }

    public class DataSplitResult
    {
        public Dictionary<string, object?[,]> SheetsToCreate { get; } = new Dictionary<string, object?[,]>(StringComparer.OrdinalIgnoreCase);
        public int TotalRowsProcessed { get; set; }
        public int DistinctGroupsCount => SheetsToCreate.Count;
        public List<string> Warnings { get; } = new List<string>();
    }

    public static class DataSplitter
    {
        private static readonly char[] ForbiddenChars = new[] { '\\', '/', '?', '*', ':', '[', ']' };

        public static DataSplitResult Split(object?[,] matrix, DataSplitOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var result = new DataSplitResult();

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int totalRows = rEnd - rStart + 1;
            int totalCols = cEnd - cStart + 1;

            if (totalRows == 0 || totalCols == 0) return result;

            int targetCol = cStart + options.SplitColumnIndex;
            if (targetCol > cEnd || targetCol < cStart)
            {
                throw new ArgumentOutOfRangeException(nameof(options.SplitColumnIndex), "Selected split column is outside range bounds.");
            }

            int dataStartRow = rStart;
            object?[]? headerRow = null;

            if (options.HasHeaders)
            {
                headerRow = new object?[totalCols];
                for (int c = 0; c < totalCols; c++)
                {
                    headerRow[c] = matrix[rStart, cStart + c];
                }
                dataStartRow++;
            }

            var comparer = options.CaseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            var groups = new Dictionary<string, List<object?[]>>(comparer);

            int rowsProcessed = 0;
            for (int r = dataStartRow; r <= rEnd; r++)
            {
                rowsProcessed++;
                object? keyVal = matrix[r, targetCol];
                string keyStr = keyVal?.ToString() ?? "";
                if (options.TrimKeyValues) keyStr = keyStr.Trim();
                if (string.IsNullOrEmpty(keyStr)) keyStr = "(Blank)";

                if (!groups.TryGetValue(keyStr, out var rowList))
                {
                    rowList = new List<object?[]>();
                    groups[keyStr] = rowList;
                }

                object?[] rowValues = new object?[totalCols];
                for (int c = 0; c < totalCols; c++)
                {
                    rowValues[c] = matrix[r, cStart + c];
                }
                rowList.Add(rowValues);
            }

            result.TotalRowsProcessed = rowsProcessed;

            // Generate clean sheet matrices
            var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var kvp in groups)
            {
                string rawName = string.IsNullOrEmpty(options.SheetNamePrefix) 
                    ? kvp.Key 
                    : $"{options.SheetNamePrefix}{kvp.Key}";

                string sanitizedName = SanitizeSheetName(rawName, usedSheetNames);
                usedSheetNames.Add(sanitizedName);

                var rowList = kvp.Value;
                int outputRowCount = (options.HasHeaders ? 1 : 0) + rowList.Count;
                object?[,] sheetMatrix = new object?[outputRowCount, totalCols];

                int curRow = 0;
                if (options.HasHeaders && headerRow != null)
                {
                    for (int c = 0; c < totalCols; c++)
                    {
                        sheetMatrix[0, c] = headerRow[c];
                    }
                    curRow = 1;
                }

                for (int i = 0; i < rowList.Count; i++)
                {
                    var values = rowList[i];
                    for (int c = 0; c < totalCols; c++)
                    {
                        sheetMatrix[curRow + i, c] = values[c];
                    }
                }

                result.SheetsToCreate[sanitizedName] = sheetMatrix;
            }

            return result;
        }

        public static string SanitizeSheetName(string candidate, HashSet<string>? existingNames = null)
        {
            if (string.IsNullOrWhiteSpace(candidate)) candidate = "Sheet";

            // Strip forbidden characters
            string clean = candidate;
            foreach (char c in ForbiddenChars)
            {
                clean = clean.Replace(c, '_');
            }

            // Remove leading/trailing apostrophes
            clean = clean.Trim().Trim('\'');

            // Handle reserved name
            if (string.Equals(clean, "History", StringComparison.OrdinalIgnoreCase))
            {
                clean = "History_Data";
            }

            // Truncate to 31 characters max
            if (clean.Length > 31)
            {
                clean = clean.Substring(0, 31).Trim();
            }

            if (string.IsNullOrEmpty(clean)) clean = "Data";

            // Uniquify
            if (existingNames != null && existingNames.Contains(clean))
            {
                string baseName = clean.Length > 27 ? clean.Substring(0, 27) : clean;
                int counter = 1;
                string uniqueCandidate;
                do
                {
                    uniqueCandidate = $"{baseName}_{counter++}";
                } while (existingNames.Contains(uniqueCandidate));

                clean = uniqueCandidate;
            }

            return clean;
        }
    }
}
