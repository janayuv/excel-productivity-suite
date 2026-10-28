using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Abstractions;

namespace ExcelProductivitySuite.UnitTests
{
    public class MockExcelContext : IExcelContext
    {
        public IntPtr MainWindowHandle => IntPtr.Zero;
        public bool HasActiveWorkbook => true;
        public bool HasActiveWorksheet => true;
        public string? ActiveWorkbookPath { get; set; } = @"C:\Test\SalesData.xlsx";
        public string? ActiveWorkbookName { get; set; } = "SalesData.xlsx";
        public string? ActiveWorksheetName { get; set; } = "Sheet1";
        public string? SelectionAddress { get; set; } = "$A$1:$B$10";
        public int SelectedCellCount => 20;

        public int SelectionStartRow { get; set; } = 1;
        public int SelectionStartColumn { get; set; } = 1;
        public int SelectionRowCount { get; set; } = 10;
        public int SelectionColumnCount { get; set; } = 2;

        public bool IsCurrentSheetProtected { get; set; } = false;
        public bool IsWorkbookProtected { get; set; } = false;
        public bool HasMergedCellsInSelection { get; set; } = false;
        public bool IsMultiAreaSelection { get; set; } = false;

        public bool ConfirmResponse { get; set; } = true;
        public bool HasDataResponse { get; set; } = false;
        public (int row, int col, int rCount, int cCount)? LastSelectedRange { get; set; }

        public object?[,]? CurrentValues { get; set; }
        public List<string> MessagesShown { get; } = new();
        public List<string> Sheets { get; } = new() { "Sheet1", "Sheet2", "Sheet3" };
        public Dictionary<string, bool> SheetVisibility { get; } = new()
        {
            { "Sheet1", true },
            { "Sheet2", false },
            { "Sheet3", true }
        };

        public bool CheckRangeHasData(int startRow, int startCol, int rowCount, int colCount) => HasDataResponse;

        public object?[,]? GetRangeValues(int startRow, int startCol, int rowCount, int colCount) => CurrentValues;

        public object?[,]? GetRangeValuesByAddress(string address) => CurrentValues;

        public void SelectRange(int startRow, int startCol, int rowCount, int colCount)
        {
            LastSelectedRange = (startRow, startCol, rowCount, colCount);
        }

        public List<(int row, int col)> LastSelectedCells { get; } = new();
        public List<(int row, int col, int color)> LastHighlightedCells { get; } = new();
        public List<CommentInfo> MockComments { get; } = new();

        public void SelectCells(IEnumerable<(int row, int col)> cells)
        {
            LastSelectedCells.Clear();
            if (cells != null) LastSelectedCells.AddRange(cells);
        }

        public void HighlightCells(IEnumerable<(int row, int col)> cells, int colorRgb)
        {
            LastHighlightedCells.Clear();
            if (cells != null)
            {
                foreach (var c in cells)
                {
                    LastHighlightedCells.Add((c.row, c.col, colorRgb));
                }
            }
        }

        public IReadOnlyList<CommentInfo> GetWorksheetComments(string? sheetName = null) => MockComments;

        public bool Confirm(string message, string title) => ConfirmResponse;

        public object?[,]? GetSelectedValues()
        {
            return CurrentValues != null ? (object?[,])CurrentValues.Clone() : null;
        }

        public List<string> GetSampleOriginalValues(int maxSamples = 10)
        {
            var samples = new List<string>();
            if (CurrentValues == null) return samples;

            int rStart = CurrentValues.GetLowerBound(0);
            int rEnd = CurrentValues.GetUpperBound(0);
            int cStart = CurrentValues.GetLowerBound(1);

            for (int r = rStart; r <= rEnd && samples.Count < maxSamples; r++)
            {
                samples.Add(CurrentValues[r, cStart]?.ToString() ?? "");
            }
            return samples;
        }

        public void SetSelectedValues(object?[,] values)
        {
            CurrentValues = (object?[,])values.Clone();
        }

        public void SetRangeValues(int startRow, int startCol, object?[,] values)
        {
            CurrentValues = (object?[,])values.Clone();
        }

        public void SetRangeValuesByAddress(string address, object?[,] values)
        {
            CurrentValues = (object?[,])values.Clone();
        }

        public IReadOnlyList<string> GetSelectedAreaAddresses()
        {
            return new List<string> { SelectionAddress ?? "$A$1:$B$10" };
        }

        public void ClearSelection()
        {
            if (CurrentValues != null)
            {
                int rStart = CurrentValues.GetLowerBound(0);
                int rEnd = CurrentValues.GetUpperBound(0);
                int cStart = CurrentValues.GetLowerBound(1);
                int cEnd = CurrentValues.GetUpperBound(1);
                for (int r = rStart; r <= rEnd; r++)
                    for (int c = cStart; c <= cEnd; c++)
                        CurrentValues[r, c] = null;
            }
        }

        public int ConvertSelectionFormulasToValues()
        {
            return SelectedCellCount;
        }

        public IReadOnlyList<string> GetWorksheetNames() => Sheets;

        public IReadOnlyList<SheetVisibilityInfo> GetWorksheetVisibilities()
        {
            var list = new List<SheetVisibilityInfo>();
            foreach (var s in Sheets)
            {
                bool isVis = SheetVisibility.TryGetValue(s, out var v) && v;
                list.Add(new SheetVisibilityInfo(s, isVis, false));
            }
            return list;
        }

        public void SetWorksheetVisibility(string sheetName, bool visible, bool veryHidden = false)
        {
            SheetVisibility[sheetName] = visible;
        }

        public void RenameWorksheets(IDictionary<string, string> renameMap)
        {
            foreach (var pair in renameMap)
            {
                int idx = Sheets.IndexOf(pair.Key);
                if (idx >= 0)
                {
                    Sheets[idx] = pair.Value;
                }
            }
        }

        public Dictionary<string, object?[,]> CreatedSheets { get; } = new();
        public List<string> OpenWorkbooks { get; } = new() { "Workbook1.xlsx", "Workbook2.xlsx" };

        public bool CreateWorksheetWithData(string sheetName, object?[,] data, bool activate = true)
        {
            CreatedSheets[sheetName] = data;
            Sheets.Add(sheetName);
            return true;
        }

        public IReadOnlyList<string> GetOpenWorkbookNames() => OpenWorkbooks;

        public IReadOnlyList<string> GetWorkbookSheetNames(string workbookName) => Sheets;

        public object?[,]? GetWorkbookRangeValues(string workbookName, string sheetName, string? rangeAddress = null) => CurrentValues;

        public Dictionary<string, object?[,]> SheetDataMap { get; } = new();
        public List<string> BlankSheets { get; } = new();

        public object?[,]? GetActiveWorksheetRangeValues(string sheetName, string? rangeAddress = null)
        {
            if (SheetDataMap.TryGetValue(sheetName, out var data)) return data;
            return CurrentValues;
        }

        public void ReorderWorksheets(IList<string> orderedSheetNames)
        {
            Sheets.Clear();
            Sheets.AddRange(orderedSheetNames);
        }

        public IReadOnlyList<string> GetBlankWorksheetNames() => BlankSheets;

        public int DeleteWorksheets(IEnumerable<string> sheetNames)
        {
            int deleted = 0;
            foreach (var s in sheetNames)
            {
                if (Sheets.Remove(s)) deleted++;
            }
            return deleted;
        }

        public bool CreateIndexWorksheet(string indexSheetName, IEnumerable<string> sheetNames, bool includeBackLinks = false)
        {
            Sheets.Insert(0, indexSheetName);
            return true;
        }

        public int SplitWorksheetsToFiles(IEnumerable<string> sheetNames, string targetDirectory, SplitFileFormat format, string namingPattern)
        {
            int count = 0;
            foreach (var s in sheetNames) count++;
            return count;
        }

        public int ImportSheetsFromWorkbook(string filePath, bool firstSheetOnly, bool prefixWithFileName)
        {
            string importedSheet = prefixWithFileName ? $"Imported_Sheet1" : "Imported_Sheet";
            Sheets.Add(importedSheet);
            return 1;
        }

        public List<CellColorInfo> MockCellColors { get; } = new();

        public IReadOnlyList<CellColorInfo> GetSelectedCellColors(int maxCells = 5000) => MockCellColors;

        public bool ApplyPreventDuplicateValidation(string? customErrorMessage = null) => true;

        public object?[,]? CurrentFormulas { get; set; }
        public List<FormulaErrorInfo> MockFormulaErrors { get; } = new();
        public List<(string src, List<string> targets, PageSetupOptions opts)> CopiedPageSetups { get; } = new();
        public (List<string> targets, FreezePanesSyncMode mode, int row, int col)? LastSynchronizedFreezePanes { get; set; }

        public object?[,]? GetSelectedFormulas() => CurrentFormulas != null ? (object?[,])CurrentFormulas.Clone() : null;

        public void SetSelectedFormulas(object?[,] formulas)
        {
            CurrentFormulas = (object?[,])formulas.Clone();
        }

        public object?[,]? GetRangeFormulas(int startRow, int startCol, int rowCount, int colCount) => CurrentFormulas;

        public object?[,]? GetRangeFormulasByAddress(string address) => CurrentFormulas;

        public void SetRangeFormulas(int startRow, int startCol, object?[,] formulas)
        {
            CurrentFormulas = (object?[,])formulas.Clone();
        }

        public void SetRangeFormulasByAddress(string address, object?[,] formulas)
        {
            CurrentFormulas = (object?[,])formulas.Clone();
        }

        public IReadOnlyList<FormulaErrorInfo> GetFormulaErrors(bool selectionOnly = false) => MockFormulaErrors;

        public int CopyPageSetup(string sourceSheetName, IEnumerable<string> targetSheetNames, PageSetupOptions options)
        {
            var targets = new List<string>(targetSheetNames);
            CopiedPageSetups.Add((sourceSheetName, targets, options));
            return targets.Count;
        }

        public int SynchronizeFreezePanes(IEnumerable<string> targetSheetNames, FreezePanesSyncMode mode, int splitRow = 1, int splitCol = 1)
        {
            var targets = new List<string>(targetSheetNames);
            LastSynchronizedFreezePanes = (targets, mode, splitRow, splitCol);
            return targets.Count;
        }

        public List<(string dest, List<string>? sheets, bool open)> ExportedPdfs { get; } = new();
        public List<(List<string> files, string outDir, string ext)> ConvertedWorkbooks { get; } = new();

        public bool ExportToPdf(string destinationPath, IEnumerable<string>? sheetNames = null, bool openAfterExport = false)
        {
            ExportedPdfs.Add((destinationPath, sheetNames != null ? new List<string>(sheetNames) : null, openAfterExport));
            return true;
        }

        public int BatchConvertWorkbooks(IEnumerable<string> filePaths, string outputDirectory, string targetExtension)
        {
            var files = new List<string>(filePaths);
            ConvertedWorkbooks.Add((files, outputDirectory, targetExtension));
            return files.Count;
        }

        public List<SheetProtectionStatus> MockProtectionStatuses { get; } = new();
        public List<(List<string> sheets, string? pwd, SheetProtectionOptions opts)> ProtectedSheetsHistory { get; } = new();
        public List<(List<string> sheets, string? pwd)> UnprotectedSheetsHistory { get; } = new();
        public List<string> SavedCopies { get; } = new();
        public string? RestoredSnapshotPath { get; set; }
        public (string title, int sRow, int sCol, int rCount, int cCount, int catCol, int actCol, int tgtCol)? LastCreatedChart { get; set; }

        public IReadOnlyList<SheetProtectionStatus> GetSheetProtectionStatuses()
        {
            if (MockProtectionStatuses.Count > 0) return MockProtectionStatuses;
            var list = new List<SheetProtectionStatus>();
            foreach (var s in Sheets)
            {
                list.Add(new SheetProtectionStatus(s, false));
            }
            return list;
        }

        public (int successCount, List<string> errors) BatchProtectWorksheets(IEnumerable<string> sheetNames, string? password, SheetProtectionOptions options)
        {
            var list = new List<string>(sheetNames);
            ProtectedSheetsHistory.Add((list, password, options));
            return (list.Count, new List<string>());
        }

        public (int successCount, List<string> errors) BatchUnprotectWorksheets(IEnumerable<string> sheetNames, string? password)
        {
            var list = new List<string>(sheetNames);
            UnprotectedSheetsHistory.Add((list, password));
            return (list.Count, new List<string>());
        }

        public bool SaveActiveWorkbookCopy(string destinationPath)
        {
            SavedCopies.Add(destinationPath);
            return true;
        }

        public bool RestoreWorkbookFromSnapshot(string snapshotPath)
        {
            RestoredSnapshotPath = snapshotPath;
            return true;
        }

        public bool CreateTargetVsActualChart(string title, int startRow, int startCol, int rowCount, int colCount, int categoryColIndex, int actualColIndex, int targetColIndex)
        {
            LastCreatedChart = (title, startRow, startCol, rowCount, colCount, categoryColIndex, actualColIndex, targetColIndex);
            return true;
        }

        public List<NamedRangeInfo> MockNamedRanges { get; } = new();
        public string? LastNavigatedNamedRange { get; set; }
        public string? LastDeletedNamedRange { get; set; }
        public List<WorksheetColumnInfo> MockColumns { get; } = new();
        public int? LastSelectedColumn { get; set; }
        public List<(List<int> cols, bool hidden)> ColumnVisibilityHistory { get; } = new();
        public List<List<int>> AutoFittedColumnsHistory { get; } = new();
        public CellContentDetails? MockCellDetails { get; set; }
        public (int row, int col, string rowHeader, string colHeader, string address) MockCellContext { get; set; } = (1, 1, "ID", "101", "$A$1");
        public (int row, int col)? LastHighlightedRowAndCol { get; set; }
        public TableDataModel? MockTableData { get; set; }
        public List<(object?[] row, int sRow, int sCol, int cCount)> AppendedRowsHistory { get; } = new();

        public IReadOnlyList<NamedRangeInfo> GetNamedRanges()
        {
            if (MockNamedRanges.Count > 0) return MockNamedRanges;
            return new List<NamedRangeInfo>
            {
                new NamedRangeInfo("SalesTotal", "=Sheet1!$B$10", "Workbook", "Total Q1 Sales", true, "Sheet1!$B$10"),
                new NamedRangeInfo("TaxRates", "=Sheet2!$A$1:$B$5", "Sheet2", "Regional tax table", true, "Sheet2!$A$1:$B$5"),
                new NamedRangeInfo("BrokenRef", "=#REF!", "Workbook", "", false, "#REF!")
            };
        }

        public bool GoToNamedRange(string name)
        {
            LastNavigatedNamedRange = name;
            return true;
        }

        public bool DeleteNamedRange(string name)
        {
            LastDeletedNamedRange = name;
            return true;
        }

        public IReadOnlyList<WorksheetColumnInfo> GetWorksheetColumns(int maxCols = 100)
        {
            if (MockColumns.Count > 0) return MockColumns;
            return new List<WorksheetColumnInfo>
            {
                new WorksheetColumnInfo(1, "A", "ID", 10.0, false),
                new WorksheetColumnInfo(2, "B", "Name", 25.0, false),
                new WorksheetColumnInfo(3, "C", "Department", 20.0, true),
                new WorksheetColumnInfo(4, "D", "Salary", 15.0, false)
            };
        }

        public void SelectColumn(int columnIndex)
        {
            LastSelectedColumn = columnIndex;
        }

        public void SetColumnsVisibility(IEnumerable<int> columnIndices, bool isHidden)
        {
            ColumnVisibilityHistory.Add((new List<int>(columnIndices), isHidden));
        }

        public void AutoFitColumns(IEnumerable<int> columnIndices)
        {
            AutoFittedColumnsHistory.Add(new List<int>(columnIndices));
        }

        public CellContentDetails? GetActiveCellDetails()
        {
            if (MockCellDetails != null) return MockCellDetails;
            return new CellContentDetails("$B$5", "Sheet1", "Sample unclipped cell content\nSecond line text", "=SUM(A1:A4)", "$#,##0.00", "Numeric");
        }

        public (int row, int col, string rowHeader, string colHeader, string address) GetActiveCellContext() => MockCellContext;

        public void HighlightActiveRowAndColumn(int row, int col)
        {
            LastHighlightedRowAndCol = (row, col);
        }

        public TableDataModel? GetActiveTableData(int maxSampleRows = 500)
        {
            if (MockTableData != null) return MockTableData;
            var headers = new List<string> { "ID", "Name", "Department", "Salary" };
            var rows = new List<object?[]>
            {
                new object?[] { 101, "Alice", "Finance", 85000 },
                new object?[] { 102, "Bob", "Engineering", 92000 }
            };
            return new TableDataModel(headers, rows, 1, 1);
        }

        public bool AppendTableRow(object?[] rowData, int tableStartRow, int tableStartCol, int colCount)
        {
            AppendedRowsHistory.Add((rowData, tableStartRow, tableStartCol, colCount));
            return true;
        }

        public void ShowMessage(string message, string title, bool isError = false)
        {
            MessagesShown.Add($"[{title}] {message}");
        }
    }
}
