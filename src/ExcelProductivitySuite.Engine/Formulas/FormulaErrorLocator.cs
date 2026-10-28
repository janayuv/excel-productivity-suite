using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Abstractions;

namespace ExcelProductivitySuite.Engine.Formulas
{
    /// <summary>
    /// Engine for detecting, categorizing, and locating formula errors across worksheets and ranges.
    /// Handles both Excel COM CVErr integer error codes and text error representations.
    /// </summary>
    public static class FormulaErrorLocator
    {
        // Standard Excel COM CVErr values
        public const int CVErrDiv0 = -2146826281;  // 0x800A07D7
        public const int CVErrNA = -2146826246;    // 0x800A07FA
        public const int CVErrName = -2146826288;  // 0x800A07D0
        public const int CVErrNull = -2146826282;  // 0x800A07D6
        public const int CVErrNum = -2146826252;   // 0x800A07F4
        public const int CVErrRef = -2146826273;   // 0x800A07DF
        public const int CVErrValue = -2146826259; // 0x800A07ED

        /// <summary>
        /// Identifies whether a cell value represents an Excel calculation or formula error.
        /// </summary>
        public static bool IsError(object? value, out string errorText)
        {
            errorText = string.Empty;
            if (value == null) return false;

            if (value is int intVal)
            {
                switch (intVal)
                {
                    case CVErrDiv0:
                        errorText = "#DIV/0!";
                        return true;
                    case CVErrNA:
                        errorText = "#N/A";
                        return true;
                    case CVErrName:
                        errorText = "#NAME?";
                        return true;
                    case CVErrNull:
                        errorText = "#NULL!";
                        return true;
                    case CVErrNum:
                        errorText = "#NUM!";
                        return true;
                    case CVErrRef:
                        errorText = "#REF!";
                        return true;
                    case CVErrValue:
                        errorText = "#VALUE!";
                        return true;
                }
            }

            string str = value.ToString()?.Trim() ?? string.Empty;
            if (str.StartsWith("#"))
            {
                if (str.Equals("#DIV/0!", StringComparison.OrdinalIgnoreCase) ||
                    str.Equals("#N/A", StringComparison.OrdinalIgnoreCase) ||
                    str.Equals("#VALUE!", StringComparison.OrdinalIgnoreCase) ||
                    str.Equals("#REF!", StringComparison.OrdinalIgnoreCase) ||
                    str.Equals("#NAME?", StringComparison.OrdinalIgnoreCase) ||
                    str.Equals("#NUM!", StringComparison.OrdinalIgnoreCase) ||
                    str.Equals("#NULL!", StringComparison.OrdinalIgnoreCase) ||
                    str.Equals("#GETTING_DATA", StringComparison.OrdinalIgnoreCase))
                {
                    errorText = str.ToUpperInvariant();
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Scans values and formulas matrices to locate all error cells with their addresses and formulas.
        /// </summary>
        public static List<FormulaErrorInfo> LocateErrors(
            object?[,] values,
            object?[,]? formulas,
            int startRow,
            int startCol,
            string sheetName)
        {
            var errors = new List<FormulaErrorInfo>();
            if (values == null) return errors;

            int vRowStart = values.GetLowerBound(0);
            int vRowEnd = values.GetUpperBound(0);
            int vColStart = values.GetLowerBound(1);
            int vColEnd = values.GetUpperBound(1);

            int? fRowStart = formulas?.GetLowerBound(0);
            int? fColStart = formulas?.GetLowerBound(1);

            for (int r = vRowStart; r <= vRowEnd; r++)
            {
                for (int c = vColStart; c <= vColEnd; c++)
                {
                    object? val = values[r, c];
                    if (IsError(val, out string errorText))
                    {
                        int actualRow = startRow + (r - vRowStart);
                        int actualCol = startCol + (c - vColStart);
                        string cellAddr = FormatA1Address(actualRow, actualCol);

                        string formulaStr = string.Empty;
                        if (formulas != null && fRowStart.HasValue && fColStart.HasValue)
                        {
                            int fR = fRowStart.Value + (r - vRowStart);
                            int fC = fColStart.Value + (c - vColStart);
                            if (fR <= formulas.GetUpperBound(0) && fC <= formulas.GetUpperBound(1))
                            {
                                formulaStr = formulas[fR, fC]?.ToString() ?? string.Empty;
                            }
                        }

                        errors.Add(new FormulaErrorInfo(sheetName, cellAddr, actualRow, actualCol, formulaStr, errorText));
                    }
                }
            }

            return errors;
        }

        public static string FormatA1Address(int row, int col)
        {
            return $"${GetColumnLetter(col)}${row}";
        }

        public static string GetColumnLetter(int col)
        {
            if (col < 1) return "A";
            string colLetter = string.Empty;
            while (col > 0)
            {
                int mod = (col - 1) % 26;
                colLetter = (char)('A' + mod) + colLetter;
                col = (col - mod) / 26;
            }
            return colLetter;
        }
    }
}
