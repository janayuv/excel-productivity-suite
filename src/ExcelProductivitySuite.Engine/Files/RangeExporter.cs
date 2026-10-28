using System;
using System.IO;
using System.Text;

namespace ExcelProductivitySuite.Engine.Files
{
    public enum ExportFileFormat
    {
        Csv,
        Tsv,
        CustomDelimited,
        Json
    }

    public class RangeExportOptions
    {
        public ExportFileFormat Format { get; set; } = ExportFileFormat.Csv;
        public string Delimiter { get; set; } = ",";
        public bool HasHeaders { get; set; } = true;
        public bool QuoteAllFields { get; set; } = false;
        public Encoding Encoding { get; set; } = Encoding.UTF8;
    }

    /// <summary>
    /// Engine for exporting 2D cell matrices to CSV, TSV, custom-delimited text, or structured JSON.
    /// </summary>
    public static class RangeExporter
    {
        public static void ExportToFile(object?[,] matrix, string filePath, RangeExportOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            options ??= new RangeExportOptions();

            string exported = ExportToString(matrix, options);
            File.WriteAllText(filePath, exported, options.Encoding);
        }

        public static string ExportToString(object?[,] matrix, RangeExportOptions? options = null)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            options ??= new RangeExportOptions();

            if (options.Format == ExportFileFormat.Json)
            {
                return ExportToJson(matrix, options.HasHeaders);
            }

            string delimiter = options.Format switch
            {
                ExportFileFormat.Tsv => "\t",
                ExportFileFormat.CustomDelimited => options.Delimiter,
                _ => ","
            };

            return ExportToDelimited(matrix, delimiter, options.QuoteAllFields);
        }

        private static string ExportToDelimited(object?[,] matrix, string delimiter, bool quoteAll)
        {
            var sb = new StringBuilder();

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    if (c > cStart) sb.Append(delimiter);

                    object? cell = matrix[r, c];
                    string cellStr = cell?.ToString() ?? string.Empty;

                    bool needsQuote = quoteAll ||
                                      cellStr.Contains(delimiter) ||
                                      cellStr.Contains("\"") ||
                                      cellStr.Contains("\r") ||
                                      cellStr.Contains("\n");

                    if (needsQuote)
                    {
                        sb.Append('"');
                        sb.Append(cellStr.Replace("\"", "\"\""));
                        sb.Append('"');
                    }
                    else
                    {
                        sb.Append(cellStr);
                    }
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private static string ExportToJson(object?[,] matrix, bool hasHeaders)
        {
            var sb = new StringBuilder();

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);
            int cols = cEnd - cStart + 1;

            if (hasHeaders && rEnd > rStart)
            {
                // Export as array of JSON objects keyed by header names
                string[] headers = new string[cols];
                for (int c = 0; c < cols; c++)
                {
                    headers[c] = matrix[rStart, cStart + c]?.ToString() ?? $"Column{c + 1}";
                }

                sb.AppendLine("[");
                for (int r = rStart + 1; r <= rEnd; r++)
                {
                    sb.Append("  {");
                    for (int c = 0; c < cols; c++)
                    {
                        if (c > 0) sb.Append(", ");
                        sb.Append('\"').Append(EscapeJson(headers[c])).Append("\": ");

                        object? val = matrix[r, cStart + c];
                        AppendJsonValue(sb, val);
                    }
                    sb.Append("}");
                    if (r < rEnd) sb.Append(",");
                    sb.AppendLine();
                }
                sb.AppendLine("]");
            }
            else
            {
                // Export as array of arrays
                sb.AppendLine("[");
                for (int r = rStart; r <= rEnd; r++)
                {
                    sb.Append("  [");
                    for (int c = 0; c < cols; c++)
                    {
                        if (c > 0) sb.Append(", ");
                        object? val = matrix[r, cStart + c];
                        AppendJsonValue(sb, val);
                    }
                    sb.Append("]");
                    if (r < rEnd) sb.Append(",");
                    sb.AppendLine();
                }
                sb.AppendLine("]");
            }

            return sb.ToString();
        }

        private static void AppendJsonValue(StringBuilder sb, object? val)
        {
            if (val == null)
            {
                sb.Append("null");
            }
            else if (val is int || val is long || val is short)
            {
                sb.Append(val);
            }
            else if (val is double d)
            {
                sb.Append(d.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            else if (val is bool b)
            {
                sb.Append(b ? "true" : "false");
            }
            else
            {
                sb.Append('\"').Append(EscapeJson(val.ToString() ?? "")).Append('\"');
            }
        }

        private static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("\\", "\\\\")
                    .Replace("\"", "\\\"")
                    .Replace("\r", "\\r")
                    .Replace("\n", "\\n")
                    .Replace("\t", "\\t");
        }
    }
}
