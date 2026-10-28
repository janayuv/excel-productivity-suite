using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum CombineDirection
    {
        AcrossColumns, // Row by row
        DownRows,      // Column by column
        EntireSelection // All into one single cell
    }

    public class CombineOptions
    {
        public CombineDirection Direction { get; set; } = CombineDirection.AcrossColumns;
        public string Delimiter { get; set; } = ", ";
        public bool SkipBlankCells { get; set; } = true;
        public bool TrimTokens { get; set; } = true;
        public string Prefix { get; set; } = "";
        public string Suffix { get; set; } = "";
    }

    public static class TextCombiner
    {
        public static string CombineList(IEnumerable<string?> items, CombineOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (items == null) return string.Empty;

            var tokens = new List<string>();
            foreach (var item in items)
            {
                if (item == null)
                {
                    if (!options.SkipBlankCells) tokens.Add(string.Empty);
                    continue;
                }

                string str = options.TrimTokens ? item.Trim() : item;
                if (string.IsNullOrEmpty(str))
                {
                    if (!options.SkipBlankCells) tokens.Add(string.Empty);
                    continue;
                }

                tokens.Add(str);
            }

            if (tokens.Count == 0) return string.Empty;

            string combined = string.Join(options.Delimiter ?? "", tokens);
            if (!string.IsNullOrEmpty(options.Prefix)) combined = options.Prefix + combined;
            if (!string.IsNullOrEmpty(options.Suffix)) combined = combined + options.Suffix;

            return combined;
        }

        public static object?[,] CombineMatrix(object?[,] matrix, CombineOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rowCount = rEnd - rStart + 1;
            int colCount = cEnd - cStart + 1;

            if (options.Direction == CombineDirection.AcrossColumns)
            {
                object?[,] result = new object?[rowCount, 1];
                for (int r = 0; r < rowCount; r++)
                {
                    var rowItems = new List<string?>();
                    for (int c = 0; c < colCount; c++)
                    {
                        var val = matrix[rStart + r, cStart + c];
                        rowItems.Add(val?.ToString());
                    }
                    result[r, 0] = CombineList(rowItems, options);
                }
                return result;
            }
            else if (options.Direction == CombineDirection.DownRows)
            {
                object?[,] result = new object?[1, colCount];
                for (int c = 0; c < colCount; c++)
                {
                    var colItems = new List<string?>();
                    for (int r = 0; r < rowCount; r++)
                    {
                        var val = matrix[rStart + r, cStart + c];
                        colItems.Add(val?.ToString());
                    }
                    result[0, c] = CombineList(colItems, options);
                }
                return result;
            }
            else // EntireSelection
            {
                object?[,] result = new object?[1, 1];
                var allItems = new List<string?>();
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        var val = matrix[rStart + r, cStart + c];
                        allItems.Add(val?.ToString());
                    }
                }
                result[0, 0] = CombineList(allItems, options);
                return result;
            }
        }
    }
}
