using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelProductivitySuite.Engine.Files
{
    public class CsvImportOptions
    {
        public string Delimiter { get; set; } = ",";
        public bool HasHeaderRow { get; set; } = true;
        public Encoding FileEncoding { get; set; } = Encoding.UTF8;
        public bool TrimWhitespace { get; set; } = true;
        public bool ParseNumbers { get; set; } = true;
    }

    /// <summary>
    /// Engine for parsing CSV and delimited text files into 2D object matrices for Excel.
    /// Complies with RFC 4180 CSV specifications, handling embedded delimiters, escaped quotes, and newlines.
    /// </summary>
    public static class CsvTextImporter
    {
        public static object?[,] ParseFile(string filePath, CsvImportOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException("File not found", filePath);

            options ??= new CsvImportOptions();
            string content = File.ReadAllText(filePath, options.FileEncoding);
            return ParseText(content, options);
        }

        public static List<List<string>> PreviewFirstRows(string filePath, CsvImportOptions? options = null, int maxRows = 10)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return new List<List<string>>();

            options ??= new CsvImportOptions();
            var result = new List<List<string>>();

            using var reader = new StreamReader(filePath, options.FileEncoding);
            var parsedRows = ParseLinesFromReader(reader, options.Delimiter, maxRows);
            return parsedRows;
        }

        public static object?[,] ParseText(string content, CsvImportOptions? options = null)
        {
            if (content == null) throw new ArgumentNullException(nameof(content));
            options ??= new CsvImportOptions();

            using var reader = new StringReader(content);
            var parsedRows = ParseLinesFromReader(reader, options.Delimiter, int.MaxValue);

            if (parsedRows.Count == 0) return new object?[0, 0];

            int rowCount = parsedRows.Count;
            int colCount = 0;
            foreach (var r in parsedRows)
            {
                if (r.Count > colCount) colCount = r.Count;
            }

            object?[,] matrix = new object?[rowCount, colCount];

            for (int r = 0; r < rowCount; r++)
            {
                var rowList = parsedRows[r];
                for (int c = 0; c < colCount; c++)
                {
                    if (c < rowList.Count)
                    {
                        string cellText = rowList[c];
                        if (options.TrimWhitespace) cellText = cellText.Trim();

                        if (options.ParseNumbers && double.TryParse(cellText, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double num))
                        {
                            matrix[r, c] = num;
                        }
                        else
                        {
                            matrix[r, c] = cellText;
                        }
                    }
                    else
                    {
                        matrix[r, c] = string.Empty;
                    }
                }
            }

            return matrix;
        }

        private static List<List<string>> ParseLinesFromReader(TextReader reader, string delimiter, int maxRows)
        {
            var rows = new List<List<string>>();
            var currentRow = new List<string>();
            var currentField = new StringBuilder();

            bool inQuotes = false;
            char delimChar = delimiter.Length == 1 ? delimiter[0] : ',';

            int chInt;
            while ((chInt = reader.Read()) != -1)
            {
                char ch = (char)chInt;

                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        int nextCh = reader.Peek();
                        if (nextCh == '"')
                        {
                            // Escaped quote
                            reader.Read();
                            currentField.Append('"');
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        currentField.Append(ch);
                    }
                }
                else
                {
                    if (ch == '"')
                    {
                        inQuotes = true;
                    }
                    else if (ch == delimChar)
                    {
                        currentRow.Add(currentField.ToString());
                        currentField.Clear();
                    }
                    else if (ch == '\r')
                    {
                        if (reader.Peek() == '\n') reader.Read();
                        currentRow.Add(currentField.ToString());
                        currentField.Clear();
                        rows.Add(currentRow);
                        currentRow = new List<string>();

                        if (rows.Count >= maxRows) break;
                    }
                    else if (ch == '\n')
                    {
                        currentRow.Add(currentField.ToString());
                        currentField.Clear();
                        rows.Add(currentRow);
                        currentRow = new List<string>();

                        if (rows.Count >= maxRows) break;
                    }
                    else
                    {
                        currentField.Append(ch);
                    }
                }
            }

            if (currentField.Length > 0 || currentRow.Count > 0)
            {
                currentRow.Add(currentField.ToString());
                rows.Add(currentRow);
            }

            return rows;
        }
    }
}
