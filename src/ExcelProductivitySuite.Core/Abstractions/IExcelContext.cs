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
        /// Gets the full file system path to the active workbook, or null if unsaved or no workbook is open.
        /// </summary>
        string? ActiveWorkbookPath { get; }

        /// <summary>
        /// Gets the filename of the active workbook, or null if no workbook is open.
        /// </summary>
        string? ActiveWorkbookName { get; }

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
        /// Writes values into a range identified by address (e.g. "$A$1" or "Sheet2!$C$5:$F$10").
        /// </summary>
        void SetRangeValuesByAddress(string address, object?[,] values);

        /// <summary>
        /// Gets the individual addresses of all discontinuous areas in a multi-area selection.
        /// </summary>
        IReadOnlyList<string> GetSelectedAreaAddresses();

        /// <summary>
        /// Clears all cell contents from the current selection.
        /// </summary>
        void ClearSelection();

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
        /// Creates a new worksheet with the specified data in a single batch assignment.
        /// Returns true if creation succeeded.
        /// </summary>
        bool CreateWorksheetWithData(string sheetName, object?[,] data, bool activate = true);

        /// <summary>
        /// Gets the names of all currently open workbooks in Excel.
        /// </summary>
        IReadOnlyList<string> GetOpenWorkbookNames();

        /// <summary>
        /// Gets the sheet names of a specific open workbook.
        /// </summary>
        IReadOnlyList<string> GetWorkbookSheetNames(string workbookName);

        /// <summary>
        /// Reads range or UsedRange values from a specific open workbook and sheet.
        /// </summary>
        object?[,]? GetWorkbookRangeValues(string workbookName, string sheetName, string? rangeAddress = null);

        /// <summary>
        /// Selects disparate individual cells on the active worksheet.
        /// </summary>
        void SelectCells(IEnumerable<(int row, int col)> cells);

        /// <summary>
        /// Applies a background fill color to specified cells on the active worksheet.
        /// </summary>
        void HighlightCells(IEnumerable<(int row, int col)> cells, int colorRgb);

        /// <summary>
        /// Retrieves all cell comments from the specified worksheet (or active sheet if null).
        /// </summary>
        IReadOnlyList<CommentInfo> GetWorksheetComments(string? sheetName = null);

        /// <summary>
        /// Reads range or UsedRange values from a specific worksheet in the active workbook.
        /// </summary>
        object?[,]? GetActiveWorksheetRangeValues(string sheetName, string? rangeAddress = null);

        /// <summary>
        /// Reorders worksheet tabs in the active workbook to match the specified sequence.
        /// </summary>
        void ReorderWorksheets(IList<string> orderedSheetNames);

        /// <summary>
        /// Detects and returns names of all completely blank worksheets in the active workbook.
        /// </summary>
        IReadOnlyList<string> GetBlankWorksheetNames();

        /// <summary>
        /// Deletes the specified worksheets from the active workbook.
        /// Returns the count of successfully deleted worksheets.
        /// </summary>
        int DeleteWorksheets(IEnumerable<string> sheetNames);

        /// <summary>
        /// Creates an interactive Table of Contents worksheet with clickable hyperlinks to target sheets.
        /// </summary>
        bool CreateIndexWorksheet(string indexSheetName, IEnumerable<string> sheetNames, bool includeBackLinks = false);

        /// <summary>
        /// Splits worksheets of the active workbook into separate individual files.
        /// Returns the count of files successfully written.
        /// </summary>
        int SplitWorksheetsToFiles(IEnumerable<string> sheetNames, string targetDirectory, SplitFileFormat format, string namingPattern);

        /// <summary>
        /// Imports worksheets from an external workbook file into the active workbook.
        /// Returns the count of imported sheets.
        /// </summary>
        int ImportSheetsFromWorkbook(string filePath, bool firstSheetOnly, bool prefixWithFileName);

        /// <summary>
        /// Reads color information (fill and font color) and values for cells in the current selection.
        /// </summary>
        IReadOnlyList<CellColorInfo> GetSelectedCellColors(int maxCells = 5000);

        /// <summary>
        /// Applies an Excel Data Validation rule to prevent duplicate entries within the selection.
        /// </summary>
        bool ApplyPreventDuplicateValidation(string? customErrorMessage = null);

        /// <summary>
        /// Reads formula strings from the current selection into a 2D object array.
        /// </summary>
        object?[,]? GetSelectedFormulas();

        /// <summary>
        /// Writes formula strings back into the current selection in a single batch assignment.
        /// </summary>
        void SetSelectedFormulas(object?[,] formulas);

        /// <summary>
        /// Reads formula strings from an arbitrary range into a 2D object array.
        /// </summary>
        object?[,]? GetRangeFormulas(int startRow, int startCol, int rowCount, int colCount);

        /// <summary>
        /// Reads formula strings from a specified range address into a 2D object array.
        /// </summary>
        object?[,]? GetRangeFormulasByAddress(string address);

        /// <summary>
        /// Writes formula strings into a target range specified by row and column coordinates.
        /// </summary>
        void SetRangeFormulas(int startRow, int startCol, object?[,] formulas);

        /// <summary>
        /// Writes formula strings into a target range specified by an A1-style address.
        /// </summary>
        void SetRangeFormulasByAddress(string address, object?[,] formulas);

        /// <summary>
        /// Locates and audits formula error cells in the current worksheet or active selection.
        /// </summary>
        IReadOnlyList<FormulaErrorInfo> GetFormulaErrors(bool selectionOnly = false);

        /// <summary>
        /// Copies page setup configurations from a source worksheet to one or more target worksheets.
        /// Returns the number of worksheets successfully updated.
        /// </summary>
        int CopyPageSetup(string sourceSheetName, IEnumerable<string> targetSheetNames, PageSetupOptions options);

        /// <summary>
        /// Synchronizes freeze panes settings across multiple target worksheets.
        /// Returns the number of worksheets successfully synchronized.
        /// </summary>
        int SynchronizeFreezePanes(IEnumerable<string> targetSheetNames, FreezePanesSyncMode mode, int splitRow = 1, int splitCol = 1);

        /// <summary>
        /// Exports specified worksheets or the entire active workbook to a PDF document.
        /// </summary>
        bool ExportToPdf(string destinationPath, IEnumerable<string>? sheetNames = null, bool openAfterExport = false);

        /// <summary>
        /// Batch converts workbook files from one format to another (e.g. .xls to .xlsx).
        /// Returns the number of files successfully converted.
        /// </summary>
        int BatchConvertWorkbooks(IEnumerable<string> filePaths, string outputDirectory, string targetExtension);

        /// <summary>
        /// Gets the protection statuses of all worksheets in the active workbook.
        /// </summary>
        IReadOnlyList<SheetProtectionStatus> GetSheetProtectionStatuses();

        /// <summary>
        /// Protects the specified worksheets with optional password and configurable permission flags.
        /// Returns the number of successfully protected sheets and any error messages.
        /// </summary>
        (int successCount, List<string> errors) BatchProtectWorksheets(IEnumerable<string> sheetNames, string? password, SheetProtectionOptions options);

        /// <summary>
        /// Unprotects the specified worksheets using the supplied password.
        /// Returns the number of successfully unprotected sheets and any error messages.
        /// </summary>
        (int successCount, List<string> errors) BatchUnprotectWorksheets(IEnumerable<string> sheetNames, string? password);

        /// <summary>
        /// Saves a non-disruptive copy of the active workbook to the specified target path.
        /// </summary>
        bool SaveActiveWorkbookCopy(string destinationPath);

        /// <summary>
        /// Closes the current active workbook without saving and opens the specified snapshot workbook.
        /// </summary>
        bool RestoreWorkbookFromSnapshot(string snapshotPath);

        /// <summary>
        /// Creates a Target vs. Actual combo chart (Clustered Column + Line with Markers) for the specified data range.
        /// </summary>
        bool CreateTargetVsActualChart(string title, int startRow, int startCol, int rowCount, int colCount, int categoryColIndex, int actualColIndex, int targetColIndex);

        /// <summary>
        /// Shows an informational, warning, or error alert dialog in Excel.
        /// </summary>
        void ShowMessage(string message, string title, bool isError = false);
    }

    public class FormulaErrorInfo
    {
        public string SheetName { get; }
        public string CellAddress { get; }
        public int Row { get; }
        public int Column { get; }
        public string Formula { get; }
        public string ErrorText { get; }

        public FormulaErrorInfo(string sheetName, string cellAddress, int row, int col, string formula, string errorText)
        {
            SheetName = sheetName;
            CellAddress = cellAddress;
            Row = row;
            Column = col;
            Formula = formula;
            ErrorText = errorText;
        }
    }

    public class PageSetupOptions
    {
        public bool CopyOrientationAndSize { get; set; } = true;
        public bool CopyMargins { get; set; } = true;
        public bool CopyPrintTitles { get; set; } = true;
        public bool CopyScaling { get; set; } = true;
        public bool CopyHeadersFooters { get; set; } = true;
        public bool CopySheetOptions { get; set; } = true;
    }

    public enum FreezePanesSyncMode
    {
        FreezeTopRow,
        FreezeFirstColumn,
        FreezeAtActiveCell,
        UnfreezeAll
    }

    public class CellColorInfo
    {
        public int Row { get; }
        public int Column { get; }
        public int FillColorRgb { get; }
        public int FontColorRgb { get; }
        public object? Value { get; }

        public CellColorInfo(int row, int col, int fillColorRgb, int fontColorRgb, object? value)
        {
            Row = row;
            Column = col;
            FillColorRgb = fillColorRgb;
            FontColorRgb = fontColorRgb;
            Value = value;
        }
    }

    public enum SplitFileFormat
    {
        ExcelWorkbook,
        Pdf,
        Csv
    }

    public class CommentInfo
    {
        public string SheetName { get; }
        public string CellAddress { get; }
        public int Row { get; }
        public int Column { get; }
        public string Author { get; }
        public string Text { get; }

        public CommentInfo(string sheetName, string cellAddress, int row, int col, string author, string text)
        {
            SheetName = sheetName;
            CellAddress = cellAddress;
            Row = row;
            Column = col;
            Author = author;
            Text = text;
        }
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

    public class SheetProtectionStatus
    {
        public string SheetName { get; }
        public bool IsProtected { get; }

        public SheetProtectionStatus(string sheetName, bool isProtected)
        {
            SheetName = sheetName;
            IsProtected = isProtected;
        }
    }

    public class SheetProtectionOptions
    {
        public bool SelectLockedCells { get; set; } = true;
        public bool SelectUnlockedCells { get; set; } = true;
        public bool FormatCells { get; set; } = false;
        public bool FormatColumns { get; set; } = false;
        public bool FormatRows { get; set; } = false;
        public bool InsertColumns { get; set; } = false;
        public bool InsertRows { get; set; } = false;
        public bool InsertHyperlinks { get; set; } = false;
        public bool DeleteColumns { get; set; } = false;
        public bool DeleteRows { get; set; } = false;
        public bool Sort { get; set; } = false;
        public bool AutoFilter { get; set; } = false;
        public bool UsePivotTables { get; set; } = false;
    }
}
