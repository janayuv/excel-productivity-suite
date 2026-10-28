using System;

namespace ExcelProductivitySuite.Engine.Text
{
    public enum PositionRemovalMode
    {
        FromLeft,
        FromRight,
        FromMiddle,
        BeforeDelimiter,
        AfterDelimiter
    }

    public class PositionRemovalOptions
    {
        public PositionRemovalMode Mode { get; set; } = PositionRemovalMode.FromLeft;
        public int Count { get; set; } = 1;
        public int StartIndex { get; set; } = 1; // 1-based index
        public string Delimiter { get; set; } = "-";
        public bool SkipBlankCells { get; set; } = true;
    }

    public static class PositionRemover
    {
        public static string Remove(string input, PositionRemovalOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrEmpty(input)) return string.Empty;

            switch (options.Mode)
            {
                case PositionRemovalMode.FromLeft:
                    int removeLeft = Math.Max(0, options.Count);
                    if (removeLeft >= input.Length) return string.Empty;
                    return input.Substring(removeLeft);

                case PositionRemovalMode.FromRight:
                    int removeRight = Math.Max(0, options.Count);
                    if (removeRight >= input.Length) return string.Empty;
                    return input.Substring(0, input.Length - removeRight);

                case PositionRemovalMode.FromMiddle:
                    int startIdx = Math.Max(0, options.StartIndex - 1); // 0-based
                    if (startIdx >= input.Length) return input;
                    int count = Math.Max(0, options.Count);
                    int remaining = input.Length - startIdx;
                    int removeCount = Math.Min(count, remaining);
                    return input.Remove(startIdx, removeCount);

                case PositionRemovalMode.BeforeDelimiter:
                    if (string.IsNullOrEmpty(options.Delimiter)) return input;
                    int idxBefore = input.IndexOf(options.Delimiter, StringComparison.OrdinalIgnoreCase);
                    if (idxBefore < 0) return input;
                    return input.Substring(idxBefore + options.Delimiter.Length);

                case PositionRemovalMode.AfterDelimiter:
                    if (string.IsNullOrEmpty(options.Delimiter)) return input;
                    int idxAfter = input.IndexOf(options.Delimiter, StringComparison.OrdinalIgnoreCase);
                    if (idxAfter < 0) return input;
                    return input.Substring(0, idxAfter);

                default:
                    return input;
            }
        }

        public static object?[,] RemoveRange(object?[,] matrix, PositionRemovalOptions options)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            var result = (object?[,])matrix.Clone();

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    var val = matrix[r, c];
                    if (val != null)
                    {
                        string str = val.ToString() ?? "";
                        if (!string.IsNullOrEmpty(str))
                        {
                            result[r, c] = Remove(str, options);
                        }
                    }
                }
            }

            return result;
        }
    }
}
