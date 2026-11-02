using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Models;

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
        /// Retrieves comprehensive metadata for all currently open workbooks.
        /// </summary>
        IReadOnlyList<WorkbookSummaryInfo> GetOpenWorkbookDetails();

        /// <summary>
        /// Retrieves comprehensive metadata for all worksheets in the specified (or active) workbook.
        /// </summary>
        IReadOnlyList<WorksheetSummaryInfo> GetWorksheetDetails(string? workbookName = null);

        /// <summary>
        /// Activates the specified open workbook window in Excel.
        /// </summary>
        bool ActivateWorkbook(string workbookName);

        /// <summary>
        /// Closes the specified open workbook, optionally saving changes.
        /// </summary>
        bool CloseWorkbook(string workbookName, bool saveChanges = false);

        /// <summary>
        /// Saves the specified open workbook.
        /// </summary>
        bool SaveWorkbook(string workbookName);

        /// <summary>
        /// Saves all currently open workbooks that have unsaved changes.
        /// </summary>
        void SaveAllWorkbooks();

        /// <summary>
        /// Creates and activates a new blank workbook in Excel.
        /// </summary>
        bool CreateNewWorkbook();

        /// <summary>
        /// Activates the specified worksheet in the specified (or active) workbook.
        /// </summary>
        bool ActivateWorksheet(string sheetName, string? workbookName = null);

        /// <summary>
        /// Moves a worksheet up, down, to top, or to end in its workbook tab order.
        /// </summary>
        bool MoveWorksheet(string sheetName, SheetMoveDirection direction, string? workbookName = null);

        /// <summary>
        /// Toggles back to the previously active worksheet in the active workbook.
        /// </summary>
        bool ToggleLastActiveWorksheet(string? workbookName = null);

        /// <summary>
        /// Adds a new blank worksheet to the active (or specified) workbook.
        /// </summary>
        bool AddWorksheet(string? desiredName = null, string? workbookName = null);

        /// <summary>
        /// Deletes the specified worksheet from the active (or specified) workbook.
        /// </summary>
        bool DeleteWorksheet(string sheetName, string? workbookName = null);

        /// <summary>
        /// Renames a single worksheet in the active (or specified) workbook.
        /// </summary>
        bool RenameWorksheet(string oldName, string newName, string? workbookName = null);

        /// <summary>
        /// Shows or toggles the docked Workbook & Sheet Navigation Task Pane in Excel, or opens the navigation dialog.
        /// </summary>
        void ShowNavigationPane();

        /// <summary>
        /// Gets or sets the active Reading View / Layout configuration settings.
        /// </summary>
        ReadingViewSettings ReadingViewSettings { get; set; }

        /// <summary>
        /// Gets whether the live Reading View crosshair highlighter is currently active.
        /// </summary>
        bool IsReadingViewActive { get; }

        /// <summary>
        /// Toggles Reading View live crosshair highlight on or off.
        /// </summary>
        void ToggleReadingView();

        /// <summary>
        /// Enables live Reading View crosshair highlighting.
        /// </summary>
        void EnableReadingView(ReadingViewSettings? settings = null);

        /// <summary>
        /// Disables live Reading View crosshair highlighting and removes all overlay shapes.
        /// </summary>
        void DisableReadingView();

        /// <summary>
        /// Updates the Reading View crosshair highlight overlay to match the active cell position.
        /// </summary>
        void UpdateReadingViewHighlight();

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
        /// Exports worksheets, selection, or entire workbook to PDF with comprehensive layout, quality, and output options.
        /// </summary>
        PdfExportResult ExportToPdfWithOptions(PdfExportSettings settings);

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
        /// Retrieves all named ranges defined in the active workbook and its worksheets.
        /// </summary>
        IReadOnlyList<NamedRangeInfo> GetNamedRanges();

        /// <summary>
        /// Navigates and selects the specified named range.
        /// </summary>
        bool GoToNamedRange(string name);

        /// <summary>
        /// Deletes the specified named range from the active workbook.
        /// </summary>
        bool DeleteNamedRange(string name);

        /// <summary>
        /// Retrieves column information (letter, header, width, visibility) for the active worksheet.
        /// </summary>
        IReadOnlyList<WorksheetColumnInfo> GetWorksheetColumns(int maxCols = 100);

        /// <summary>
        /// Selects the specified column on the active worksheet.
        /// </summary>
        void SelectColumn(int columnIndex);

        /// <summary>
        /// Sets the visibility (hidden / visible) of the specified columns.
        /// </summary>
        void SetColumnsVisibility(IEnumerable<int> columnIndices, bool isHidden);

        /// <summary>
        /// Auto-fits column widths for the specified columns.
        /// </summary>
        void AutoFitColumns(IEnumerable<int> columnIndices);

        /// <summary>
        /// Retrieves rich inspection details for the current active cell (text, formula, formatting, metrics).
        /// </summary>
        CellContentDetails? GetActiveCellDetails();

        /// <summary>
        /// Retrieves header and coordinate context for the current active cell (for grid focus).
        /// </summary>
        (int row, int col, string rowHeader, string colHeader, string address) GetActiveCellContext();

        /// <summary>
        /// Selects the entire row and column of the given coordinates (crosshair focus).
        /// </summary>
        void HighlightActiveRowAndColumn(int row, int col);

        /// <summary>
        /// Retrieves structured tabular data from the active selection or UsedRange for a dynamic form.
        /// </summary>
        TableDataModel? GetActiveTableData(int maxSampleRows = 500);

        /// <summary>
        /// Appends a new data row to the bottom of the active table or dataset.
        /// </summary>
        bool AppendTableRow(object?[] rowData, int tableStartRow, int tableStartCol, int colCount);

        /// <summary>
        /// Shows an informational, warning, or error alert dialog in Excel.
        /// </summary>
        void ShowMessage(string message, string title, bool isError = false);

        /// <summary>
        /// Retrieves all precedent cells contributing to the active cell formula.
        /// </summary>
        IReadOnlyList<CellReferenceInfo> GetActiveCellPrecedents(bool directOnly = true);

        /// <summary>
        /// Retrieves all cells dependent upon the active cell value.
        /// </summary>
        IReadOnlyList<CellReferenceInfo> GetActiveCellDependents(bool directOnly = true);

        /// <summary>
        /// Navigates and selects the specified cell address on a worksheet.
        /// </summary>
        bool SelectCellAddress(string sheetName, string cellAddress);

        /// <summary>
        /// Applies an Excel NumberFormat string across the active selection or worksheet.
        /// </summary>
        void ApplyNumberFormat(string formatString, bool selectionOnly = true);

        /// <summary>
        /// Gets the number format mask of the active cell.
        /// </summary>
        string? GetActiveCellNumberFormat();

        /// <summary>
        /// Inserts a formula expression directly into the active cell cursor.
        /// </summary>
        void InsertFormulaAtActiveCell(string formula);

        /// <summary>
        /// Retrieves 2D matrix values from all disjoint areas in the active multi-selection.
        /// </summary>
        IReadOnlyList<object?[,]> GetMultiAreaSelectedValues();

        /// <summary>
        /// Writes matrix values back to each disjoint area in the active multi-selection.
        /// </summary>
        void SetMultiAreaSelectedValues(IReadOnlyList<object?[,]> areaValues);

        /// <summary>
        /// Retrieves cell addresses of each disjoint area in the active multi-selection.
        /// </summary>
        IReadOnlyList<string> GetMultiAreaAddresses();

        /// <summary>
        /// Writes a 2D matrix directly to a specified destination range address.
        /// </summary>
        bool WriteMatrixToAddress(string targetAddress, object?[,] matrix);

        /// <summary>
        /// Retrieves row indices of visible rows within a specified row range, skipping hidden/filtered rows.
        /// </summary>
        IReadOnlyList<int> GetVisibleRowIndices(int startRow, int maxRows = 2000);

        /// <summary>
        /// Pastes matrix data strictly into visible cells starting at the given coordinates, bypassing hidden rows.
        /// </summary>
        bool PasteIntoVisibleCells(object?[,] sourceMatrix, int startRow, int startCol);

        /// <summary>
        /// Retrieves information on all merged cell areas within the current selection.
        /// </summary>
        IReadOnlyList<MergedAreaInfo> GetMergedAreasInSelection();

        /// <summary>
        /// Unmerges all composite cells across the selection and populates all constituent cells with the original value.
        /// </summary>
        int UnmergeAndFillSelection();

        /// <summary>
        /// Creates a new worksheet and writes the given matrix data into it.
        /// </summary>
        bool WriteMatrixToNewWorksheet(string sheetName, object?[,] matrix);

        /// <summary>
        /// Scans worksheet or selection for cells matching specific font, fill, or style criteria.
        /// </summary>
        IReadOnlyList<FormatMatchItem> FindCellsByFormat(CellFormatCriteria criteria, bool selectionOnly = false);

        /// <summary>
        /// Reads format attributes (fill color, font color, bold, italic) from the current active cell cursor.
        /// </summary>
        CellFormatCriteria GetActiveCellFormatCriteria();

        /// <summary>
        /// Audits data validation rules across the worksheet or active selection.
        /// </summary>
        IReadOnlyList<ValidationRuleInfo> GetWorksheetValidationRules(bool selectionOnly = false);

        /// <summary>
        /// Selects multiple cells across the worksheet given their A1 addresses.
        /// </summary>
        void SelectCellsByAddresses(IEnumerable<string> addresses);

        /// <summary>
        /// Removes data validation rules from the specified cell addresses.
        /// </summary>
        void RemoveValidationFromRange(IEnumerable<string> addresses);

        /// <summary>
        /// Clears data validation rules across the specified scope.
        /// </summary>
        int ClearValidationRules(ValidationClearScope scope);

        /// <summary>
        /// Creates an in-cell dropdown list data validation rule on the selected range.
        /// </summary>
        bool CreateDropdownList(string listFormulaOrItems, bool isFormula, bool allowBlank, bool showDropdown, string? errorTitle, string? errorMessage);

        /// <summary>
        /// Adjusts row heights and/or column widths across the selection or active worksheet.
        /// </summary>
        bool AdjustCellDimensions(double? rowHeightPoints, double? colWidthPoints, bool autoFitRows, bool autoFitCols, bool entireSheet);

        /// <summary>
        /// Inserts customizable sequential numbers, IDs, or codes into the active selection or designated cells.
        /// Supports contiguous ranges, non-contiguous multi-area selections, and merged cells.
        /// </summary>
        OperationResult InsertSequenceNumbers(SequenceNumberDefinition sequence, SequenceFillOptions options);

        /// <summary>
        /// Transforms columns to tables, tables to single columns/rows, or reshapes matrices based on fixed or delimiter rules.
        /// </summary>
        OperationResult TransformRange(TransformRangeOptions options);

        /// <summary>
        /// Central undo/redo manager tracking cell and range modifications across suite operations.
        /// </summary>
        IUndoManager UndoManager { get; }

        /// <summary>
        /// Reverts the most recent suite command that modified cell values or formulas.
        /// </summary>
        bool UndoLastAction();

        /// <summary>
        /// Re-applies the most recent undone suite command.
        /// </summary>
        bool RedoLastAction();
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

    public class NamedRangeInfo
    {
        public string Name { get; }
        public string RefersTo { get; }
        public string Scope { get; }
        public string Comment { get; }
        public bool IsValid { get; }
        public string TargetAddress { get; }

        public NamedRangeInfo(string name, string refersTo, string scope, string comment, bool isValid, string targetAddress)
        {
            Name = name;
            RefersTo = refersTo;
            Scope = scope;
            Comment = comment;
            IsValid = isValid;
            TargetAddress = targetAddress;
        }
    }

    public class WorksheetColumnInfo
    {
        public int Index { get; }
        public string ColumnLetter { get; }
        public string Header { get; }
        public double Width { get; }
        public bool IsHidden { get; }
        public string DisplayName => $"{ColumnLetter}: {(string.IsNullOrWhiteSpace(Header) ? "(Empty)" : Header)}";

        public WorksheetColumnInfo(int index, string columnLetter, string header, double width, bool isHidden)
        {
            Index = index;
            ColumnLetter = columnLetter;
            Header = header;
            Width = width;
            IsHidden = isHidden;
        }
    }

    public class CellContentDetails
    {
        public string Address { get; }
        public string SheetName { get; }
        public string ValueText { get; }
        public string FormulaText { get; }
        public string NumberFormat { get; }
        public int CharCount { get; }
        public int WordCount { get; }
        public int LineCount { get; }
        public string DataType { get; }

        public CellContentDetails(string address, string sheetName, string valueText, string formulaText, string numberFormat, string dataType)
        {
            Address = address;
            SheetName = sheetName;
            ValueText = valueText;
            FormulaText = formulaText;
            NumberFormat = numberFormat;
            DataType = dataType;
            CharCount = valueText.Length;
            WordCount = string.IsNullOrWhiteSpace(valueText) ? 0 : valueText.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
            LineCount = string.IsNullOrWhiteSpace(valueText) ? 0 : valueText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).Length;
        }
    }

    public class TableDataModel
    {
        public IReadOnlyList<string> Headers { get; }
        public List<object?[]> Rows { get; }
        public int StartRow { get; }
        public int StartCol { get; }
        public int RowCount => Rows.Count;
        public int ColumnCount => Headers.Count;

        public TableDataModel(IReadOnlyList<string> headers, List<object?[]> rows, int startRow, int startCol)
        {
            Headers = headers;
            Rows = rows;
            StartRow = startRow;
            StartCol = startCol;
        }
    }

    public class CellReferenceInfo
    {
        public string SheetName { get; }
        public string CellAddress { get; }
        public string FullAddress => string.IsNullOrWhiteSpace(SheetName) ? CellAddress : $"'{SheetName}'!{CellAddress}";
        public string ValueText { get; }
        public string FormulaText { get; }
        public bool IsFormula => !string.IsNullOrWhiteSpace(FormulaText) && FormulaText.StartsWith("=");
        public bool IsDirect { get; }

        public CellReferenceInfo(string sheetName, string cellAddress, string valueText, string formulaText, bool isDirect = true)
        {
            SheetName = sheetName ?? string.Empty;
            CellAddress = cellAddress ?? string.Empty;
            ValueText = valueText ?? string.Empty;
            FormulaText = formulaText ?? string.Empty;
            IsDirect = isDirect;
        }
    }

    public class MergedAreaInfo
    {
        public int StartRow { get; }
        public int StartCol { get; }
        public int RowCount { get; }
        public int ColCount { get; }
        public string Address { get; }
        public object? Value { get; }
        public string ValueText => Value?.ToString() ?? string.Empty;

        public MergedAreaInfo(int startRow, int startCol, int rowCount, int colCount, string address, object? value)
        {
            StartRow = startRow;
            StartCol = startCol;
            RowCount = rowCount;
            ColCount = colCount;
            Address = address;
            Value = value;
        }
    }

    public class CellFormatCriteria
    {
        public bool MatchFillColor { get; set; }
        public int? FillColorRgb { get; set; }
        public bool MatchFontColor { get; set; }
        public int? FontColorRgb { get; set; }
        public bool MatchBold { get; set; }
        public bool? IsBold { get; set; }
        public bool MatchItalic { get; set; }
        public bool? IsItalic { get; set; }
    }

    public class FormatMatchItem
    {
        public int Row { get; }
        public int Column { get; }
        public string Address { get; }
        public string SheetName { get; }
        public string ValueText { get; }
        public int? FillColorRgb { get; }
        public int? FontColorRgb { get; }
        public bool IsBold { get; }
        public bool IsItalic { get; }

        public string FillHex => FillColorRgb.HasValue ? $"#{FillColorRgb.Value:X6}" : "None";
        public string FontHex => FontColorRgb.HasValue ? $"#{FontColorRgb.Value:X6}" : "Auto";

        public FormatMatchItem(int row, int col, string address, string sheetName, string valText, int? fillRgb, int? fontRgb, bool isBold, bool isItalic)
        {
            Row = row;
            Column = col;
            Address = address;
            SheetName = sheetName;
            ValueText = valText;
            FillColorRgb = fillRgb;
            FontColorRgb = fontRgb;
            IsBold = isBold;
            IsItalic = isItalic;
        }
    }

    public class ValidationRuleInfo
    {
        public string SheetName { get; }
        public string CellAddress { get; }
        public int Row { get; }
        public int Column { get; }
        public string ValidationType { get; }
        public string Operator { get; }
        public string Formula1 { get; }
        public string Formula2 { get; }
        public string CurrentValue { get; }
        public string ErrorMessage { get; }
        public bool? IsValid { get; }
        public string StatusDisplay => IsValid.HasValue ? (IsValid.Value ? "Valid" : "⚠️ Invalid") : "Active";

        public ValidationRuleInfo(
            string sheetName,
            string cellAddress,
            int row,
            int col,
            string validationType,
            string op,
            string formula1,
            string formula2,
            string currentValue,
            string errorMessage,
            bool? isValid)
        {
            SheetName = sheetName;
            CellAddress = cellAddress;
            Row = row;
            Column = col;
            ValidationType = validationType;
            Operator = op;
            Formula1 = formula1 ?? string.Empty;
            Formula2 = formula2 ?? string.Empty;
            CurrentValue = currentValue ?? string.Empty;
            ErrorMessage = errorMessage ?? string.Empty;
            IsValid = isValid;
        }
    }

    public enum ValidationClearScope
    {
        Selection,
        ActiveWorksheet,
        AllWorksheets
    }

    public enum CellDimensionUnit
    {
        Points,
        Inches,
        Centimeters
    }
}
