using System;
using System.IO;
using System.Text;

namespace ExcelProductivitySuite.Engine.Files
{
    public enum TextFileInsertMode
    {
        RowByRow,   // Split lines across successive rows
        SingleCell  // Entire file content into a single cell
    }

    /// <summary>
    /// Engine for loading external plain text or log files directly into Excel worksheets.
    /// </summary>
    public static class TextFileReader
    {
        public static object?[,] ReadFileToMatrix(
            string filePath,
            TextFileInsertMode mode = TextFileInsertMode.RowByRow,
            Encoding? encoding = null)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException("File not found", filePath);

            encoding ??= Encoding.UTF8;

            if (mode == TextFileInsertMode.SingleCell)
            {
                string text = File.ReadAllText(filePath, encoding);
                return new object?[1, 1] { { text } };
            }
            else
            {
                string[] lines = File.ReadAllLines(filePath, encoding);
                if (lines.Length == 0) return new object?[1, 1] { { string.Empty } };

                object?[,] matrix = new object?[lines.Length, 1];
                for (int i = 0; i < lines.Length; i++)
                {
                    matrix[i, 0] = lines[i];
                }
                return matrix;
            }
        }
    }
}
