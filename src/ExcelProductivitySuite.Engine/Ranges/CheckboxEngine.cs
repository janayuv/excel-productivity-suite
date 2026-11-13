using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Ranges
{
    /// <summary>
    /// Pure computational engine for batch inserting smart checkboxes.
    /// Handles range coordinate arithmetic, linked cell mapping, caption formatting,
    /// and dimension planning without Office COM dependencies.
    /// </summary>
    public static class CheckboxEngine
    {
        private static readonly Regex RangeRegex = new Regex(
            @"^(?:'?(?<sheet>[^'!]+)'?!)?\$?(?<startCol>[A-Za-z]+)\$?(?<startRow>\d+)(?::\$?(?<endCol>[A-Za-z]+)\$?(?<endRow>\d+))?$",
            RegexOptions.Compiled);

        /// <summary>
        /// Represents the execution plan for a single checkbox insertion.
        /// </summary>
        public class CheckboxItemPlan
        {
            public int Row { get; set; }
            public int Column { get; set; }
            public string CellAddress { get; set; } = string.Empty;
            public string? LinkedCellAddress { get; set; }
            public string Caption { get; set; } = string.Empty;
            public bool InitialChecked { get; set; }
            public bool ClearOriginalCell { get; set; }
            public object? NewCellValue { get; set; }
            public object? NewLinkedCellValue { get; set; }
        }

        /// <summary>
        /// Converts a 1-based column number into Excel column letters (e.g. 1 -> A, 27 -> AA).
        /// </summary>
        public static string ColumnIndexToLetter(int columnIndex)
        {
            if (columnIndex < 1) return "A";
            string columnLetter = string.Empty;
            while (columnIndex > 0)
            {
                int modulo = (columnIndex - 1) % 26;
                columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
                columnIndex = (columnIndex - modulo) / 26;
            }
            return columnLetter;
        }

        /// <summary>
        /// Converts Excel column letters into a 1-based column index (e.g. A -> 1, AA -> 27).
        /// </summary>
        public static int ColumnLetterToIndex(string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter)) return 1;
            int sum = 0;
            string clean = columnLetter.Trim().ToUpperInvariant();
            for (int i = 0; i < clean.Length; i++)
            {
                char c = clean[i];
                if (c < 'A' || c > 'Z') break;
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum > 0 ? sum : 1;
        }

        /// <summary>
        /// Parses an A1-notation cell or range address into 1-based coordinates.
        /// </summary>
        public static bool TryParseRangeAddress(
            string address,
            out string? sheetName,
            out int startRow,
            out int startCol,
            out int endRow,
            out int endCol)
        {
            sheetName = null;
            startRow = 1;
            startCol = 1;
            endRow = 1;
            endCol = 1;

            if (string.IsNullOrWhiteSpace(address))
                return false;

            var match = RangeRegex.Match(address.Trim());
            if (!match.Success)
                return false;

            if (match.Groups["sheet"].Success)
                sheetName = match.Groups["sheet"].Value;

            startCol = ColumnLetterToIndex(match.Groups["startCol"].Value);
            startRow = int.Parse(match.Groups["startRow"].Value);

            if (match.Groups["endCol"].Success && match.Groups["endRow"].Success)
            {
                endCol = ColumnLetterToIndex(match.Groups["endCol"].Value);
                endRow = int.Parse(match.Groups["endRow"].Value);
            }
            else
            {
                endCol = startCol;
                endRow = startRow;
            }

            // Normalize in case coordinates were reversed (e.g. B10:A1)
            if (startRow > endRow)
            {
                int temp = startRow;
                startRow = endRow;
                endRow = temp;
            }
            if (startCol > endCol)
            {
                int temp = startCol;
                startCol = endCol;
                endCol = temp;
            }

            return true;
        }

        /// <summary>
        /// Formats 1-based row and column coordinates into an Excel cell address (e.g. "$A$1" or "A1").
        /// </summary>
        public static string FormatCellAddress(int row, int col, bool absolute = true, string? sheetName = null)
        {
            string colLetter = ColumnIndexToLetter(col);
            string address = absolute ? $"${colLetter}${row}" : $"{colLetter}{row}";
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                return $"'{sheetName}'!{address}";
            }
            return address;
        }

        /// <summary>
        /// Computes bounding coordinates for a checkbox within a cell.
        /// </summary>
        public static void CalculateCheckboxBounds(
            double cellLeft,
            double cellTop,
            double cellWidth,
            double cellHeight,
            bool hasCaption,
            out double left,
            out double top,
            out double width,
            out double height)
        {
            const double standardBoxSize = 16.0;
            const double verticalMargin = 1.0;

            left = cellLeft + 2.0;
            top = cellTop + Math.Max(verticalMargin, (cellHeight - standardBoxSize) / 2.0);
            height = Math.Min(cellHeight - (verticalMargin * 2), 17.0);

            if (hasCaption)
            {
                width = Math.Max(cellWidth - 4.0, 60.0);
            }
            else
            {
                width = standardBoxSize;
            }
        }

        /// <summary>
        /// Generates a complete itemized plan for batch checkbox insertion.
        /// </summary>
        public static List<CheckboxItemPlan> PlanBatchInsert(
            object[,] cellValues,
            int startRow,
            int startCol,
            BatchInsertCheckboxOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            int rowCount = cellValues.GetLength(0);
            int colCount = cellValues.GetLength(1);
            var planList = new List<CheckboxItemPlan>(rowCount * colCount);

            int targetStartRow = startRow;
            int targetStartCol = startCol;
            string? targetSheetName = null;

            if (options.LinkMode == CheckboxLinkMode.OtherCells && !string.IsNullOrWhiteSpace(options.TargetLinkRangeAddress))
            {
                if (TryParseRangeAddress(options.TargetLinkRangeAddress!, out targetSheetName, out int tr, out int tc, out _, out _))
                {
                    targetStartRow = tr;
                    targetStartCol = tc;
                }
            }

            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    int currentRow = startRow + r;
                    int currentCol = startCol + c;
                    string cellAddress = FormatCellAddress(currentRow, currentCol, absolute: true);

                    object? originalVal = cellValues[r, c];
                    string originalText = originalVal?.ToString()?.Trim() ?? string.Empty;

                    var item = new CheckboxItemPlan
                    {
                        Row = currentRow,
                        Column = currentCol,
                        CellAddress = cellAddress,
                        InitialChecked = options.InitialChecked
                    };

                    // Caption
                    if (options.UseCellContentAsLabel && !string.IsNullOrEmpty(originalText))
                    {
                        item.Caption = originalText;
                        item.ClearOriginalCell = options.ClearOriginalContentWhenLabeled;
                    }
                    else
                    {
                        item.Caption = string.Empty;
                        item.ClearOriginalCell = false;
                    }

                    // Linked Cell & Values
                    switch (options.LinkMode)
                    {
                        case CheckboxLinkMode.OwnCell:
                            item.LinkedCellAddress = cellAddress;
                            if (options.DisplayLogicalValue)
                            {
                                item.NewCellValue = options.InitialChecked;
                            }
                            else if (item.ClearOriginalCell)
                            {
                                item.NewCellValue = null;
                            }
                            else
                            {
                                item.NewCellValue = originalVal;
                            }
                            break;

                        case CheckboxLinkMode.OtherCells:
                            int targetRow = targetStartRow + r;
                            int targetCol = targetStartCol + c;
                            item.LinkedCellAddress = FormatCellAddress(targetRow, targetCol, absolute: true, targetSheetName);
                            if (options.DisplayLogicalValue)
                            {
                                item.NewLinkedCellValue = options.InitialChecked;
                            }
                            if (item.ClearOriginalCell)
                            {
                                item.NewCellValue = null;
                            }
                            else
                            {
                                item.NewCellValue = originalVal;
                            }
                            break;

                        case CheckboxLinkMode.None:
                        default:
                            item.LinkedCellAddress = null;
                            if (item.ClearOriginalCell)
                            {
                                item.NewCellValue = null;
                            }
                            else
                            {
                                item.NewCellValue = originalVal;
                            }
                            break;
                    }

                    planList.Add(item);
                }
            }

            return planList;
        }
    }
}
