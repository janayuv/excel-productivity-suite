using System;
using System.Collections.Generic;
using System.Linq;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Files;
using ExcelProductivitySuite.Engine.Ranges;

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

        public List<WorkbookSummaryInfo> MockWorkbookSummaries { get; } = new()
        {
            new WorkbookSummaryInfo("Workbook1.xlsx", @"C:\Test\Workbook1.xlsx", true, true, 3),
            new WorkbookSummaryInfo("Workbook2.xlsx", @"C:\Test\Workbook2.xlsx", false, false, 2)
        };

        private string? _mockPrevSheet = null;

        public IReadOnlyList<WorkbookSummaryInfo> GetOpenWorkbookDetails() => MockWorkbookSummaries;

        public IReadOnlyList<WorksheetSummaryInfo> GetWorksheetDetails(string? workbookName = null)
        {
            var list = new List<WorksheetSummaryInfo>();
            for (int i = 0; i < Sheets.Count; i++)
            {
                string s = Sheets[i];
                bool isVis = !SheetVisibility.TryGetValue(s, out var v) || v;
                bool isActive = string.Equals(s, ActiveWorksheetName, StringComparison.OrdinalIgnoreCase);
                list.Add(new WorksheetSummaryInfo(s, workbookName ?? ActiveWorkbookName ?? "Workbook1.xlsx", i + 1, isActive, isVis ? SheetVisibilityState.Visible : SheetVisibilityState.Hidden));
            }
            return list;
        }

        public bool ActivateWorkbook(string workbookName)
        {
            ActiveWorkbookName = workbookName;
            for (int i = 0; i < MockWorkbookSummaries.Count; i++)
            {
                var wb = MockWorkbookSummaries[i];
                bool isAct = string.Equals(wb.Name, workbookName, StringComparison.OrdinalIgnoreCase);
                MockWorkbookSummaries[i] = new WorkbookSummaryInfo(wb.Name, wb.Path, isAct, wb.IsSaved, wb.SheetCount, wb.IsReadOnly);
            }
            return true;
        }

        public bool CloseWorkbook(string workbookName, bool saveChanges = false)
        {
            MockWorkbookSummaries.RemoveAll(w => string.Equals(w.Name, workbookName, StringComparison.OrdinalIgnoreCase));
            OpenWorkbooks.RemoveAll(w => string.Equals(w, workbookName, StringComparison.OrdinalIgnoreCase));
            return true;
        }

        public bool SaveWorkbook(string workbookName)
        {
            for (int i = 0; i < MockWorkbookSummaries.Count; i++)
            {
                var wb = MockWorkbookSummaries[i];
                if (string.Equals(wb.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                {
                    MockWorkbookSummaries[i] = new WorkbookSummaryInfo(wb.Name, wb.Path, wb.IsActive, true, wb.SheetCount, wb.IsReadOnly);
                    return true;
                }
            }
            return false;
        }

        public void SaveAllWorkbooks()
        {
            for (int i = 0; i < MockWorkbookSummaries.Count; i++)
            {
                var wb = MockWorkbookSummaries[i];
                MockWorkbookSummaries[i] = new WorkbookSummaryInfo(wb.Name, wb.Path, wb.IsActive, true, wb.SheetCount, wb.IsReadOnly);
            }
        }

        public bool CreateNewWorkbook()
        {
            string name = $"Book{MockWorkbookSummaries.Count + 1}.xlsx";
            MockWorkbookSummaries.Add(new WorkbookSummaryInfo(name, "[Unsaved]", true, false, 1));
            OpenWorkbooks.Add(name);
            ActivateWorkbook(name);
            return true;
        }

        public bool ActivateWorksheet(string sheetName, string? workbookName = null)
        {
            if (Sheets.Contains(sheetName))
            {
                if (ActiveWorksheetName != null && !string.Equals(ActiveWorksheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    _mockPrevSheet = ActiveWorksheetName;
                }
                ActiveWorksheetName = sheetName;
                SheetVisibility[sheetName] = true;
                return true;
            }
            return false;
        }

        public bool MoveWorksheet(string sheetName, SheetMoveDirection direction, string? workbookName = null)
        {
            int idx = Sheets.IndexOf(sheetName);
            if (idx < 0) return false;

            Sheets.RemoveAt(idx);
            switch (direction)
            {
                case SheetMoveDirection.Top:
                    Sheets.Insert(0, sheetName);
                    break;
                case SheetMoveDirection.End:
                    Sheets.Add(sheetName);
                    break;
                case SheetMoveDirection.Up:
                    int newUp = Math.Max(0, idx - 1);
                    Sheets.Insert(newUp, sheetName);
                    break;
                case SheetMoveDirection.Down:
                    int newDown = Math.Min(Sheets.Count, idx + 1);
                    Sheets.Insert(newDown, sheetName);
                    break;
            }
            return true;
        }

        public bool ToggleLastActiveWorksheet(string? workbookName = null)
        {
            if (_mockPrevSheet != null && Sheets.Contains(_mockPrevSheet))
            {
                return ActivateWorksheet(_mockPrevSheet, workbookName);
            }
            return false;
        }

        public bool AddWorksheet(string? desiredName = null, string? workbookName = null)
        {
            string name = desiredName ?? $"Sheet{Sheets.Count + 1}";
            Sheets.Add(name);
            SheetVisibility[name] = true;
            ActivateWorksheet(name, workbookName);
            return true;
        }

        public bool DeleteWorksheet(string sheetName, string? workbookName = null)
        {
            bool removed = Sheets.Remove(sheetName);
            if (removed && string.Equals(ActiveWorksheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                ActiveWorksheetName = Sheets.FirstOrDefault();
            }
            return removed;
        }

        public bool RenameWorksheet(string oldName, string newName, string? workbookName = null)
        {
            int idx = Sheets.IndexOf(oldName);
            if (idx >= 0)
            {
                Sheets[idx] = newName;
                if (SheetVisibility.TryGetValue(oldName, out var v))
                {
                    SheetVisibility.Remove(oldName);
                    SheetVisibility[newName] = v;
                }
                if (string.Equals(ActiveWorksheetName, oldName, StringComparison.OrdinalIgnoreCase))
                {
                    ActiveWorksheetName = newName;
                }
                return true;
            }
            return false;
        }

        public bool NavigationPaneShown { get; set; } = false;
        public void ShowNavigationPane()
        {
            NavigationPaneShown = true;
        }

        public ReadingViewSettings ReadingViewSettings { get; set; } = ReadingViewSettings.CreateDefault();
        public bool IsReadingViewActive { get; set; } = false;
        public bool IsGridFocusActive => IsReadingViewActive;
        public int ReadingViewHighlightUpdateCount { get; set; } = 0;

        public void ToggleReadingView()
        {
            IsReadingViewActive = !IsReadingViewActive;
            ReadingViewSettings.IsEnabled = IsReadingViewActive;
        }

        public void ToggleGridFocus() => ToggleReadingView();

        public void EnableReadingView(ReadingViewSettings? settings = null)
        {
            if (settings != null)
            {
                ReadingViewSettings = settings;
            }
            IsReadingViewActive = true;
            ReadingViewSettings.IsEnabled = true;
            UpdateReadingViewHighlight();
        }

        public void EnableGridFocus(ReadingViewSettings? settings = null) => EnableReadingView(settings);

        public void DisableReadingView()
        {
            IsReadingViewActive = false;
            ReadingViewSettings.IsEnabled = false;
        }

        public void DisableGridFocus() => DisableReadingView();

        public void UpdateReadingViewHighlight()
        {
            ReadingViewHighlightUpdateCount++;
        }

        public List<string> MockPrinters { get; } = new List<string> { "Microsoft Print to PDF", "OneNote (Desktop)", "HP LaserJet Pro" };
        public string? MockDefaultPrinter { get; set; } = "Microsoft Print to PDF";
        public PrintWorkbooksResult? LastBatchPrintResult { get; private set; }

        public IReadOnlyList<string> GetInstalledPrinterNames() => MockPrinters;
        public string? GetDefaultPrinterName() => MockDefaultPrinter;

        public PrintWorkbooksResult BatchPrintWorkbooks(
            PrintWorkbooksConfig config,
            IProgress<(int current, int total, string currentFile)>? progress = null,
            System.Threading.CancellationToken cancellationToken = default)
        {
            var res = new PrintWorkbooksResult();
            var (isValid, err) = ExcelProductivitySuite.Engine.Printing.PrintWorkbooksEngine.ValidateConfig(config);
            if (!isValid)
            {
                res.Errors.Add(err ?? "Invalid configuration");
                LastBatchPrintResult = res;
                return res;
            }

            var selected = config.SelectedFiles.Where(f => f.IsSelected).ToList();
            res.TotalWorkbooks = selected.Count;

            int count = 0;
            foreach (var item in selected)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    item.Status = PrintItemStatus.Skipped;
                    res.SkippedWorkbooks++;
                    continue;
                }
                count++;
                progress?.Report((count, selected.Count, item.FileName));

                item.Status = PrintItemStatus.Completed;
                item.SheetsPrintedCount = 2;
                res.PrintedWorkbooks++;
                res.TotalSheetsPrinted += 2;
            }
            res.ProcessedItems = selected;
            LastBatchPrintResult = res;
            return res;
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
        public List<PdfExportSettings> ExportedPdfSettings { get; } = new();
        public List<(List<string> files, string outDir, string ext)> ConvertedWorkbooks { get; } = new();

        public bool ExportToPdf(string destinationPath, IEnumerable<string>? sheetNames = null, bool openAfterExport = false)
        {
            var sheetList = sheetNames != null ? new List<string>(sheetNames) : new List<string>();
            var settings = new PdfExportSettings
            {
                DestinationPath = destinationPath,
                SelectedSheets = sheetList,
                Scope = sheetList.Count > 0 ? (sheetList.Count == 1 ? PdfExportScope.ActiveSheet : PdfExportScope.SelectedSheets) : PdfExportScope.EntireWorkbook,
                OpenAfterExport = openAfterExport
            };
            return ExportToPdfWithOptions(settings).Success;
        }

        public PdfExportResult ExportToPdfWithOptions(PdfExportSettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            ExportedPdfSettings.Add(settings);

            var validation = PdfExportEngine.ValidateSettings(settings, Sheets, !string.IsNullOrEmpty(SelectionAddress));
            if (!validation.Success)
            {
                return PdfExportResult.Failed(validation.Message);
            }

            var exportedFiles = new List<string>();
            if (settings.OutputMode == PdfOutputMode.SingleCombinedFile)
            {
                exportedFiles.Add(settings.DestinationPath);
                ExportedPdfs.Add((settings.DestinationPath, settings.SelectedSheets, settings.OpenAfterExport));
            }
            else
            {
                var sheets = settings.Scope == PdfExportScope.SelectedSheets && settings.SelectedSheets.Count > 0
                    ? settings.SelectedSheets
                    : Sheets;

                int idx = 1;
                foreach (var s in sheets)
                {
                    string fName = PdfExportEngine.ResolveFileName(settings.FileNamePattern, ActiveWorkbookName, s, idx++);
                    string path = System.IO.Path.Combine(settings.OutputDirectory, fName);
                    exportedFiles.Add(path);
                    ExportedPdfs.Add((path, new List<string> { s }, settings.OpenAfterExport));
                }
            }

            return PdfExportResult.Succeeded(exportedFiles.Count, exportedFiles, "Mock export succeeded.");
        }

        public List<(SequenceNumberDefinition seq, SequenceFillOptions opts)> InsertedSequences { get; } = new();

        public OperationResult InsertSequenceNumbers(SequenceNumberDefinition sequence, SequenceFillOptions options)
        {
            if (sequence == null) throw new ArgumentNullException(nameof(sequence));
            options ??= new SequenceFillOptions();

            InsertedSequences.Add((sequence.Clone(), options));

            if (CurrentValues != null)
            {
                CurrentValues = SequenceNumberEngine.FormatMatrix(
                    CurrentValues,
                    sequence,
                    options,
                    out int updated,
                    out int skipped,
                    out long next);

                sequence.CurrentNumber = next;
                return OperationResult.Succeeded($"Inserted {updated} sequential numbers.", updated, skipped);
            }

            int rows = SelectionRowCount > 0 ? SelectionRowCount : 10;
            int cols = SelectionColumnCount > 0 ? SelectionColumnCount : 1;
            var matrix = new object?[rows, cols];

            CurrentValues = SequenceNumberEngine.FormatMatrix(
                matrix,
                sequence,
                options,
                out int upCount,
                out int skCount,
                out long nextNum);

            sequence.CurrentNumber = nextNum;
            return OperationResult.Succeeded($"Inserted {upCount} sequential numbers.", upCount, skCount);
        }

        public OperationResult TransformRange(TransformRangeOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            var (isValid, errorMessage) = options.Validate();
            if (!isValid) return OperationResult.Failed(errorMessage ?? "Invalid transformation options.");

            var matrix = CurrentValues ?? new object?[1, 1];
            var transformed = RangeTransformer.Transform(matrix, options);
            int rows = transformed.GetLength(0);
            int cols = transformed.GetLength(1);

            if (options.DestinationMode == TransformDestinationMode.NewSheet)
            {
                CreateWorksheetWithData(options.NewSheetName, transformed);
            }
            else if (options.DestinationMode == TransformDestinationMode.TargetCell)
            {
                SetRangeValuesByAddress(options.TargetAddress, transformed);
            }
            else
            {
                CurrentValues = transformed;
            }

            int total = rows * cols;
            return OperationResult.Succeeded($"Transformed {rows}x{cols} range.", total);
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

        public List<CellReferenceInfo> MockPrecedents { get; } = new();
        public List<CellReferenceInfo> MockDependents { get; } = new();
        public (string sheetName, string cellAddress)? LastNavigatedCell { get; set; }
        public List<(string formatString, bool selectionOnly)> AppliedNumberFormats { get; } = new();
        public string? MockNumberFormat { get; set; } = "$#,##0.00";
        public List<string> InsertedFormulasHistory { get; } = new();

        public IReadOnlyList<CellReferenceInfo> GetActiveCellPrecedents(bool directOnly = true)
        {
            if (MockPrecedents.Count > 0) return MockPrecedents;
            return new List<CellReferenceInfo>
            {
                new CellReferenceInfo("Sheet1", "$A$1", "100", "", directOnly),
                new CellReferenceInfo("Sheet1", "$A$2", "200", "=B1*2", directOnly),
                new CellReferenceInfo("Sheet2", "$C$10", "500", "", directOnly)
            };
        }

        public IReadOnlyList<CellReferenceInfo> GetActiveCellDependents(bool directOnly = true)
        {
            if (MockDependents.Count > 0) return MockDependents;
            return new List<CellReferenceInfo>
            {
                new CellReferenceInfo("Sheet1", "$B$10", "300", "=SUM(A1:A5)", directOnly),
                new CellReferenceInfo("Sheet1", "$C$15", "600", "=B10*2", directOnly)
            };
        }

        public bool SelectCellAddress(string sheetName, string cellAddress)
        {
            LastNavigatedCell = (sheetName, cellAddress);
            return true;
        }

        public void ApplyNumberFormat(string formatString, bool selectionOnly = true)
        {
            AppliedNumberFormats.Add((formatString, selectionOnly));
        }

        public string? GetActiveCellNumberFormat() => MockNumberFormat;

        public void InsertFormulaAtActiveCell(string formula)
        {
            InsertedFormulasHistory.Add(formula);
        }

        public List<object?[,]> MockMultiAreaValues { get; } = new();
        public List<string> MockMultiAreaAddresses { get; } = new() { "$A$1:$A$5", "$C$1:$C$5" };
        public List<(string address, object?[,] matrix)> WrittenMatrices { get; } = new();
        public List<int> MockVisibleRows { get; } = new() { 1, 2, 5, 6, 8, 10 };
        public List<MergedAreaInfo> MockMergedAreas { get; } = new();
        public List<FormatMatchItem> MockFormatMatches { get; } = new();
        public List<ValidationRuleInfo> MockValidationRules { get; } = new();
        public List<string> SelectedCellAddressesHistory { get; } = new();
        public List<string> RemovedValidationHistory { get; } = new();
        public List<(string name, object?[,] matrix)> CreatedWorksheets { get; } = new();

        public IReadOnlyList<object?[,]> GetMultiAreaSelectedValues()
        {
            if (MockMultiAreaValues.Count > 0) return MockMultiAreaValues;
            return new List<object?[,]>
            {
                new object?[,] { { "A1" }, { "A2" }, { "A3" } },
                new object?[,] { { "C1" }, { "C2" }, { "C3" } }
            };
        }

        public List<IReadOnlyList<object?[,]>> SetMultiAreaValuesHistory { get; } = new();

        public void SetMultiAreaSelectedValues(IReadOnlyList<object?[,]> areaValues)
        {
            SetMultiAreaValuesHistory.Add(areaValues);
            MockMultiAreaValues.Clear();
            MockMultiAreaValues.AddRange(areaValues);
        }

        public IReadOnlyList<string> GetMultiAreaAddresses() => MockMultiAreaAddresses;

        public bool WriteMatrixToAddress(string targetAddress, object?[,] matrix)
        {
            WrittenMatrices.Add((targetAddress, matrix));
            return true;
        }

        public IReadOnlyList<int> GetVisibleRowIndices(int startRow, int maxRows = 2000) => MockVisibleRows;

        public bool PasteIntoVisibleCells(object?[,] sourceMatrix, int startRow, int startCol)
        {
            WrittenMatrices.Add(($"Visible_R{startRow}C{startCol}", sourceMatrix));
            return true;
        }

        public IReadOnlyList<MergedAreaInfo> GetMergedAreasInSelection()
        {
            if (MockMergedAreas.Count > 0) return MockMergedAreas;
            return new List<MergedAreaInfo>
            {
                new MergedAreaInfo(1, 1, 2, 2, "$A$1:$B$2", "MergedGroup1"),
                new MergedAreaInfo(4, 1, 3, 1, "$A$4:$A$6", "MergedGroup2")
            };
        }

        public int UnmergeAndFillSelection() => 5;

        public bool WriteMatrixToNewWorksheet(string sheetName, object?[,] matrix)
        {
            CreatedWorksheets.Add((sheetName, matrix));
            return true;
        }

        public IReadOnlyList<FormatMatchItem> FindCellsByFormat(CellFormatCriteria criteria, bool selectionOnly = false)
        {
            if (MockFormatMatches.Count > 0) return MockFormatMatches;
            return new List<FormatMatchItem>
            {
                new FormatMatchItem(2, 2, "$B$2", "Sheet1", "Match1", 0xFFFF00, 0x000000, true, false),
                new FormatMatchItem(5, 3, "$C$5", "Sheet1", "Match2", 0xFFFF00, 0x000000, true, false)
            };
        }

        public CellFormatCriteria GetActiveCellFormatCriteria()
        {
            return new CellFormatCriteria
            {
                MatchFillColor = true,
                FillColorRgb = 0xFFFF00,
                MatchBold = true,
                IsBold = true
            };
        }

        public IReadOnlyList<ValidationRuleInfo> GetWorksheetValidationRules(bool selectionOnly = false)
        {
            if (MockValidationRules.Count > 0) return MockValidationRules;
            return new List<ValidationRuleInfo>
            {
                new ValidationRuleInfo("Sheet1", "$A$2", 2, 1, "List", "", "High,Medium,Low", "", "High", "", true),
                new ValidationRuleInfo("Sheet1", "$B$2", 2, 2, "Whole Number", "Between", "1", "100", "150", "Must be 1-100", false)
            };
        }

        public void SelectCellsByAddresses(IEnumerable<string> addresses)
        {
            SelectedCellAddressesHistory.AddRange(addresses);
        }

        public void RemoveValidationFromRange(IEnumerable<string> addresses)
        {
            RemovedValidationHistory.AddRange(addresses);
        }

        public int ClearValidationRules(ValidationClearScope scope) => 1;

        public bool CreateDropdownList(string listFormulaOrItems, bool isFormula, bool allowBlank, bool showDropdown, string? errorTitle, string? errorMessage) => true;

        public bool AdjustCellDimensions(double? rowHeightPoints, double? colWidthPoints, bool autoFitRows, bool autoFitCols, bool entireSheet) => true;

        public IUndoManager UndoManager { get; } = new ExcelProductivitySuite.Engine.Recovery.UndoManager();

        public bool UndoLastAction()
        {
            if (!UndoManager.CanUndo) return false;
            var step = UndoManager.PopUndo();
            if (step != null)
            {
                CurrentValues = step.BeforeValues;
                UndoManager.PushRedo(step);
                return true;
            }
            return false;
        }

        public bool RedoLastAction()
        {
            if (!UndoManager.CanRedo) return false;
            var step = UndoManager.PopRedo();
            if (step != null)
            {
                CurrentValues = step.AfterValues;
                UndoManager.PushUndo(step);
                return true;
            }
            return false;
        }
    }
}
