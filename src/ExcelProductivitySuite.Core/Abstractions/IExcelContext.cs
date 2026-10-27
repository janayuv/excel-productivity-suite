using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Core.Abstractions
{
    /// <summary>
    /// Abstraction representing the host Excel environment. Keeps Excel COM references
    /// isolated from pure business logic and ViewModels.
    /// </summary>
    public interface IExcelContext
    {
        /// <summary>
        /// Gets the handle (HWND) of the main Excel host window for WPF dialog parenting.
        /// </summary>
        IntPtr MainWindowHandle { get; }

        /// <summary>
        /// Gets whether an active workbook is currently open in Excel.
        /// </summary>
        bool HasActiveWorkbook { get; }

        /// <summary>
        /// Gets whether an active worksheet is available.
        /// </summary>
        bool HasActiveWorksheet { get; }

        /// <summary>
        /// Gets the name of the active worksheet, or null if none.
        /// </summary>
        string? ActiveWorksheetName { get; }

        /// <summary>
        /// Gets the address in A1-style notation of the current user selection (e.g. "$A$1:$D$25").
        /// </summary>
        string? SelectionAddress { get; }

        /// <summary>
        /// Gets total number of cells in the current selection.
        /// </summary>
        int SelectedCellCount { get; }

        /// <summary>
        /// Gets the 1-based start row index of the selection on the active sheet.
        /// </summary>
        int SelectionStartRow { get; }

        /// <summary>
        /// Gets the 1-based start column index of the selection on the active sheet.
        /// </summary>
        int SelectionStartColumn { get; }

        /// <summary>
        /// Gets total rows in the current selection.
        /// </summary>
        int SelectionRowCount { get; }

        /// <summary>
        /// Gets total columns in the current selection.
        /// </summary>
        int SelectionColumnCount { get; }

        /// <summary>
        /// Gets whether the current active worksheet has protection enabled.
        /// </summary>
        bool IsCurrentSheetProtected { get; }

        /// <summary>
        /// Gets whether the active workbook has structure protection enabled.
        /// </summary>
        bool IsWorkbookProtected { get; }

        /// <summary>
        /// Gets whether any merged cells exist within the active selection.
        /// </summary>
        bool HasMergedCellsInSelection { get; }

        /// <summary>
        /// Gets whether the active selection consists of multiple discontinuous areas (e.g. via Ctrl+Click).
        /// </summary>
        bool IsMultiAreaSelection { get; }

        /// <summary>
        /// Checks if any non-empty cells exist in the specified target range.
        /// </summary>
        bool CheckRangeHasData(int startRow, int startCol, int rowCount, int colCount);

        /// <summary>
        /// Reads cell values from an arbitrary worksheet range into a 2D object array.
        /// </summary>
        object?[,]? GetRangeValues(int startRow, int startCol, int rowCount, int colCount);

        /// <summary>
        /// Reads cell values from a specified range address (e.g. "B1:B20" or "Sheet2!A1:D10") into a 2D object array.
        /// </summary>
        object?[,]? GetRangeValuesByAddress(string address);

        /// <summary>
        /// Selects a specified range of cells on the active worksheet.
        /// </summary>
        void SelectRange(int startRow, int startCol, int rowCount, int colCount);

        /// <summary>
        /// Shows a confirmation dialog with Yes/No choices. Returns true if user selected Yes.
        /// </summary>
        bool Confirm(string message, string title);

        /// <summary>
        /// Reads the selected cell values into a 2D object array in managed memory.
        /// Array indices are 1-based (Excel COM convention) or 0-based normalized.
        /// </summary>
        object?[,]? GetSelectedValues();

        /// <summary>
        /// Reads sample rows (up to maxRows) from the current selection for previewing.
        /// </summary>
        List<string> GetSampleOriginalValues(int maxSamples = 10);

        /// <summary>
        /// Writes an updated 2D array of values back to the selection in a single batch operation.
        /// </summary>
        void SetSelectedValues(object?[,] values);

        /// <summary>
        /// Writes values into a specified target range offset or expanded dimensions.
        /// </summary>
        void SetRangeValues(int startRow, int startCol, object?[,] values);

        /// <summary>
        /// Converts all formula cells in the selected range to static values.
        /// </summary>
        int ConvertSelectionFormulasToValues();

        /// <summary>
        /// Gets a list of all worksheet names in the active workbook.
        /// </summary>
        IReadOnlyList<string> GetWorksheetNames();

        /// <summary>
        /// Gets the visibility status of all worksheets (SheetName, IsVisible, IsVeryHidden).
        /// </summary>
        IReadOnlyList<SheetVisibilityInfo> GetWorksheetVisibilities();

        /// <summary>
        /// Updates the visibility of specified worksheets.
        /// </summary>
        void SetWorksheetVisibility(string sheetName, bool visible, bool veryHidden = false);

        /// <summary>
        /// Renames worksheets in bulk. Key: oldName, Value: newName.
        /// </summary>
        void RenameWorksheets(IDictionary<string, string> renameMap);

        /// <summary>
        /// Shows an informational, warning, or error alert dialog in Excel.
        /// </summary>
        void ShowMessage(string message, string title, bool isError = false);
    }

    public class SheetVisibilityInfo
    {
        public string SheetName { get; }
        public bool IsVisible { get; }
        public bool IsVeryHidden { get; }

        public SheetVisibilityInfo(string sheetName, bool isVisible, bool isVeryHidden)
        {
            SheetName = sheetName;
            IsVisible = isVisible;
            IsVeryHidden = isVeryHidden;
        }
    }
}
