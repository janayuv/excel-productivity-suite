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

        public void ShowMessage(string message, string title, bool isError = false)
        {
            MessagesShown.Add($"[{title}] {message}");
        }
    }
}
