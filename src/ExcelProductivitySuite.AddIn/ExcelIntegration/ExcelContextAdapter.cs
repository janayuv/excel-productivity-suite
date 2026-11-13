using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Threading;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Files;
using ExcelProductivitySuite.Engine.Formulas;
using ExcelProductivitySuite.Engine.Navigation;
using ExcelProductivitySuite.Engine.Printing;
using ExcelProductivitySuite.Engine.Ranges;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelProductivitySuite.AddIn.ExcelIntegration
{
    public class ExcelContextAdapter : IExcelContext
    {
        private readonly Func<Excel.Application> _appAccessor;
        private readonly IUndoManager _undoManager;
        private Excel.Application _app => _appAccessor() ?? throw new InvalidOperationException("Excel Application instance is not available.");

        public IUndoManager UndoManager => _undoManager;
        public Action? NavigationPaneToggleAction { get; set; }
        private string? _previousActiveSheetName;

        private ReadingViewSettings _readingViewSettings = ReadingLayoutEngine.LoadSettings();
        private bool _isReadingViewActive = false;

        public ReadingViewSettings ReadingViewSettings
        {
            get => _readingViewSettings;
            set => _readingViewSettings = value ?? ReadingViewSettings.CreateDefault();
        }

        public bool IsReadingViewActive => _isReadingViewActive;
        public bool IsGridFocusActive => _isReadingViewActive;

        public ExcelContextAdapter(Excel.Application app, IUndoManager? undoManager = null)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));
            _appAccessor = () => app;
            _undoManager = undoManager ?? new ExcelProductivitySuite.Engine.Recovery.UndoManager(50);
        }

        public ExcelContextAdapter(Func<Excel.Application> appAccessor, IUndoManager? undoManager = null)
        {
            _appAccessor = appAccessor ?? throw new ArgumentNullException(nameof(appAccessor));
            _undoManager = undoManager ?? new ExcelProductivitySuite.Engine.Recovery.UndoManager(50);
        }

        public IntPtr MainWindowHandle
        {
            get
            {
                try { return new IntPtr(_app.Hwnd); }
                catch { return IntPtr.Zero; }
            }
        }

        public bool HasActiveWorkbook
        {
            get
            {
                try { return _app.ActiveWorkbook != null; }
                catch { return false; }
            }
        }

        public string? ActiveWorkbookPath
        {
            get
            {
                try { return _app.ActiveWorkbook?.FullName; }
                catch { return null; }
            }
        }

        public string? ActiveWorkbookName
        {
            get
            {
                try { return _app.ActiveWorkbook?.Name; }
                catch { return null; }
            }
        }

        public bool HasActiveWorksheet
        {
            get
            {
                try { return _app.ActiveSheet is Excel.Worksheet; }
                catch { return false; }
            }
        }

        public string? ActiveWorksheetName
        {
            get
            {
                try
                {
                    if (_app.ActiveSheet is Excel.Worksheet ws)
                    {
                        return ws.Name;
                    }
                    return null;
                }
                catch { return null; }
            }
        }

        public string? SelectionAddress
        {
            get
            {
                try
                {
                    if (_app.Selection is Excel.Range range)
                    {
                        return range.get_Address();
                    }
                    return null;
                }
                catch { return null; }
            }
        }

        public int SelectedCellCount
        {
            get
            {
                try
                {
                    if (_app.Selection is Excel.Range range)
                    {
                        return range.CountLarge is long count ? (int)Math.Min(count, int.MaxValue) : range.Count;
                    }
                    return 0;
                }
                catch { return 0; }
            }
        }

        public int SelectionStartRow
        {
            get
            {
                try
                {
                    if (_app.Selection is Excel.Range range)
                    {
                        return range.Row;
                    }
                    return 1;
                }
                catch { return 1; }
            }
        }

        public int SelectionStartColumn
        {
            get
            {
                try
                {
                    if (_app.Selection is Excel.Range range)
                    {
                        return range.Column;
                    }
                    return 1;
                }
                catch { return 1; }
            }
        }

        public int SelectionRowCount
        {
            get
            {
                try
                {
                    if (_app.Selection is Excel.Range range)
                    {
                        return range.Rows.Count;
                    }
                    return 0;
                }
                catch { return 0; }
            }
        }

        public int SelectionColumnCount
        {
            get
            {
                try
                {
                    if (_app.Selection is Excel.Range range)
                    {
                        return range.Columns.Count;
                    }
                    return 0;
                }
                catch { return 0; }
            }
        }

        public bool IsCurrentSheetProtected
        {
            get
            {
                try
                {
                    if (_app.ActiveSheet is Excel.Worksheet ws)
                    {
                        return ws.ProtectContents;
                    }
                    return false;
                }
                catch { return false; }
            }
        }

        public bool IsWorkbookProtected
        {
            get
            {
                try
                {
                    if (_app.ActiveWorkbook is Excel.Workbook wb)
                    {
                        return wb.ProtectStructure;
                    }
                    return false;
                }
                catch { return false; }
            }
        }

        public bool HasMergedCellsInSelection
        {
            get
            {
                try
                {
                    if (_app.Selection is Excel.Range range)
                    {
                        object merge = range.MergeCells;
                        if (merge is bool b) return b;
                        return merge != null;
                    }
                    return false;
                }
                catch { return false; }
            }
        }

        public bool IsMultiAreaSelection
        {
            get
            {
                try
                {
                    if (_app.Selection is Excel.Range range)
                    {
                        return range.Areas.Count > 1;
                    }
                    return false;
                }
                catch { return false; }
            }
        }

        public object?[,]? GetSelectedValues()
        {
            try
            {
                if (!(_app.Selection is Excel.Range range)) return null;

                object raw = range.Value2;
                if (raw is object[,] matrix)
                {
                    return (object?[,])matrix;
                }

                // Single cell returns scalar object, normalize to 1-based 1x1 array for consistency
                object?[,] single = new object?[1, 1];
                single[0, 0] = raw;
                return single;
            }
            catch (Exception ex)
            {
                ShowMessage($"Error reading selection values: {ex.Message}", "Selection Error", isError: true);
                return null;
            }
        }

        public List<string> GetSampleOriginalValues(int maxSamples = 10)
        {
            var samples = new List<string>();
            try
            {
                if (IsMultiAreaSelection)
                {
                    var areas = GetMultiAreaSelectedValues();
                    foreach (var matrix in areas)
                    {
                        if (matrix == null) continue;
                        int rStart = matrix.GetLowerBound(0);
                        int rEnd = matrix.GetUpperBound(0);
                        int cStart = matrix.GetLowerBound(1);

                        for (int r = rStart; r <= rEnd && samples.Count < maxSamples; r++)
                        {
                            var val = matrix[r, cStart];
                            samples.Add(val?.ToString() ?? "");
                        }
                        if (samples.Count >= maxSamples) break;
                    }
                    return samples;
                }

                var singleMatrix = GetSelectedValues();
                if (singleMatrix == null) return samples;

                int rS = singleMatrix.GetLowerBound(0);
                int rE = singleMatrix.GetUpperBound(0);
                int cS = singleMatrix.GetLowerBound(1);

                for (int r = rS; r <= rE && samples.Count < maxSamples; r++)
                {
                    var val = singleMatrix[r, cS];
                    samples.Add(val?.ToString() ?? "");
                }
            }
            catch
            {
                // Fallback empty list
            }
            return samples;
        }

        public void SetSelectedValues(object?[,] values)
        {
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (!(_app.Selection is Excel.Range range)) return;

            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before modifying cell values.", "Sheet Protected", isError: true);
                return;
            }

            if (IsMultiAreaSelection)
            {
                ShowMessage("This operation does not support non-contiguous multi-area selections. Please select a single contiguous range.", "Multi-Area Selection", isError: true);
                return;
            }

            RecordBeforeRangeChange(range, values, isFormula: false, customDesc: "Cell Value Update");

            using (new ExcelStateScope(_app))
            {
                range.Value2 = values;
            }
        }

        public void SetRangeValues(int startRow, int startCol, object?[,] values)
        {
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (!(_app.ActiveSheet is Excel.Worksheet sheet)) return;

            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before modifying cell values.", "Sheet Protected", isError: true);
                return;
            }

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            if (rows <= 0 || cols <= 0) return;

            using (new ExcelStateScope(_app))
            {
                Excel.Range topLeft = sheet.Cells[startRow, startCol];
                Excel.Range bottomRight = sheet.Cells[startRow + rows - 1, startCol + cols - 1];
                Excel.Range targetRange = sheet.Range[topLeft, bottomRight];

                RecordBeforeRangeChange(targetRange, values, isFormula: false, customDesc: "Range Value Update");

                targetRange.Value2 = values;
            }
        }

        public void SetRangeValuesByAddress(string address, object?[,] values)
        {
            if (string.IsNullOrWhiteSpace(address)) return;
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before modifying cell values.", "Sheet Protected", isError: true);
                return;
            }

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            if (rows <= 0 || cols <= 0) return;

            using (new ExcelStateScope(_app))
            {
                Excel.Range target;
                if (address.Contains("!"))
                {
                    target = _app.Range[address];
                }
                else if (_app.ActiveSheet is Excel.Worksheet sheet)
                {
                    target = sheet.Range[address];
                }
                else
                {
                    target = _app.Range[address];
                }

                if (target != null)
                {
                    if (target.Rows.Count != rows || target.Columns.Count != cols)
                    {
                        target = target.Resize[rows, cols];
                    }

                    RecordBeforeRangeChange(target, values, isFormula: false, customDesc: "Address Value Update");

                    target.Value2 = values;
                }
            }
        }

        public IReadOnlyList<string> GetSelectedAreaAddresses()
        {
            var list = new List<string>();
            try
            {
                if (_app.Selection is Excel.Range range)
                {
                    foreach (Excel.Range area in range.Areas)
                    {
                        list.Add(area.get_Address());
                    }
                }
            }
            catch { }
            return list;
        }

        public void ClearSelection()
        {
            try
            {
                if (!(_app.Selection is Excel.Range range)) return;
                if (IsCurrentSheetProtected) return;
                using (new ExcelStateScope(_app))
                {
                    range.ClearContents();
                }
            }
            catch { }
        }

        public bool CheckRangeHasData(int startRow, int startCol, int rowCount, int colCount)
        {
            try
            {
                if (!(_app.ActiveSheet is Excel.Worksheet sheet)) return false;
                Excel.Range topLeft = sheet.Cells[startRow, startCol];
                Excel.Range bottomRight = sheet.Cells[startRow + rowCount - 1, startCol + colCount - 1];
                Excel.Range target = sheet.Range[topLeft, bottomRight];

                object raw = target.Value2;
                if (raw is object[,] matrix)
                {
                    int rStart = matrix.GetLowerBound(0);
                    int rEnd = matrix.GetUpperBound(0);
                    int cStart = matrix.GetLowerBound(1);
                    int cEnd = matrix.GetUpperBound(1);

                    for (int r = rStart; r <= rEnd; r++)
                    {
                        for (int c = cStart; c <= cEnd; c++)
                        {
                            var v = matrix[r, c];
                            if (v != null && !string.IsNullOrWhiteSpace(v.ToString()))
                            {
                                return true;
                            }
                        }
                    }
                    return false;
                }

                return raw != null && !string.IsNullOrWhiteSpace(raw.ToString());
            }
            catch
            {
                return false;
            }
        }

        public object?[,]? GetRangeValues(int startRow, int startCol, int rowCount, int colCount)
        {
            try
            {
                if (!(_app.ActiveSheet is Excel.Worksheet sheet)) return null;
                Excel.Range topLeft = sheet.Cells[startRow, startCol];
                Excel.Range bottomRight = sheet.Cells[startRow + rowCount - 1, startCol + colCount - 1];
                Excel.Range target = sheet.Range[topLeft, bottomRight];

                object raw = target.Value2;
                if (raw is object[,] matrix)
                {
                    return (object?[,])matrix;
                }

                object?[,] single = new object?[1, 1];
                single[0, 0] = raw;
                return single;
            }
            catch
            {
                return null;
            }
        }

        public object?[,]? GetRangeValuesByAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) return null;
            try
            {
                Excel.Range target;
                if (address.Contains("!"))
                {
                    target = _app.Range[address];
                }
                else if (_app.ActiveSheet is Excel.Worksheet sheet)
                {
                    target = sheet.Range[address];
                }
                else
                {
                    target = _app.Range[address];
                }

                if (target == null) return null;

                object raw = target.Value2;
                if (raw is object[,] matrix)
                {
                    return (object?[,])matrix;
                }

                object?[,] single = new object?[1, 1];
                single[0, 0] = raw;
                return single;
            }
            catch (Exception ex)
            {
                ShowMessage($"Unable to read range '{address}': {ex.Message}", "Range Error", isError: true);
                return null;
            }
        }

        public void SelectRange(int startRow, int startCol, int rowCount, int colCount)
        {
            try
            {
                if (!(_app.ActiveSheet is Excel.Worksheet sheet)) return;
                Excel.Range topLeft = sheet.Cells[startRow, startCol];
                Excel.Range bottomRight = sheet.Cells[startRow + rowCount - 1, startCol + colCount - 1];
                Excel.Range target = sheet.Range[topLeft, bottomRight];
                target.Select();
            }
            catch { }
        }

        public void SelectCells(IEnumerable<(int row, int col)> cells)
        {
            if (cells == null || !(_app.ActiveSheet is Excel.Worksheet sheet)) return;
            try
            {
                Excel.Range? unionRange = null;
                int count = 0;
                foreach (var (r, c) in cells)
                {
                    Excel.Range cell = sheet.Cells[r, c];
                    unionRange = unionRange == null ? cell : _app.Union(unionRange, cell);
                    count++;
                    if (count >= 1500) break;
                }
                unionRange?.Select();
            }
            catch { }
        }

        public void HighlightCells(IEnumerable<(int row, int col)> cells, int colorRgb)
        {
            if (cells == null || !(_app.ActiveSheet is Excel.Worksheet sheet)) return;
            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before highlighting.", "Sheet Protected", isError: true);
                return;
            }

            try
            {
                using (new ExcelStateScope(_app))
                {
                    Excel.Range? unionRange = null;
                    int count = 0;
                    foreach (var (r, c) in cells)
                    {
                        Excel.Range cell = sheet.Cells[r, c];
                        unionRange = unionRange == null ? cell : _app.Union(unionRange, cell);
                        count++;
                        if (count % 200 == 0 && unionRange != null)
                        {
                            unionRange.Interior.Color = colorRgb;
                            unionRange = null;
                        }
                    }
                    if (unionRange != null) unionRange.Interior.Color = colorRgb;
                }
            }
            catch { }
        }

        public IReadOnlyList<CommentInfo> GetWorksheetComments(string? sheetName = null)
        {
            var list = new List<CommentInfo>();
            try
            {
                Excel.Worksheet? sheet = null;
                if (!string.IsNullOrEmpty(sheetName))
                {
                    sheet = _app.Sheets[sheetName] as Excel.Worksheet;
                }
                else
                {
                    sheet = _app.ActiveSheet as Excel.Worksheet;
                }

                if (sheet == null) return list;

                string targetSheetName = sheet.Name;
                foreach (Excel.Comment c in sheet.Comments)
                {
                    try
                    {
                        var cell = c.Parent as Excel.Range;
                        string addr = cell?.get_Address() ?? "";
                        int row = cell?.Row ?? 0;
                        int col = cell?.Column ?? 0;
                        string author = c.Author ?? "";
                        string text = c.Text() ?? "";

                        list.Add(new CommentInfo(targetSheetName, addr, row, col, author, text));
                    }
                    catch { }
                }
            }
            catch { }
            return list;
        }

        public bool Confirm(string message, string title)
        {
            return MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes;
        }

        public int ConvertSelectionFormulasToValues()
        {
            if (!(_app.Selection is Excel.Range range)) return 0;

            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before converting formulas.", "Sheet Protected", isError: true);
                return 0;
            }

            int count = 0;
            using (new ExcelStateScope(_app))
            {
                foreach (Excel.Range area in range.Areas)
                {
                    area.Value2 = area.Value2;
                    count += area.CountLarge is long c ? (int)Math.Min(c, int.MaxValue) : area.Count;
                }
            }

            return count;
        }

        public IReadOnlyList<string> GetWorksheetNames()
        {
            var list = new List<string>();
            try
            {
                if (_app.ActiveWorkbook != null)
                {
                    foreach (Excel.Worksheet ws in _app.ActiveWorkbook.Worksheets)
                    {
                        list.Add(ws.Name);
                    }
                }
            }
            catch { }
            return list;
        }

        public IReadOnlyList<SheetVisibilityInfo> GetWorksheetVisibilities()
        {
            var list = new List<SheetVisibilityInfo>();
            try
            {
                if (_app.ActiveWorkbook != null)
                {
                    foreach (Excel.Worksheet ws in _app.ActiveWorkbook.Worksheets)
                    {
                        bool isVisible = ws.Visible == Excel.XlSheetVisibility.xlSheetVisible;
                        bool isVeryHidden = ws.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden;
                        list.Add(new SheetVisibilityInfo(ws.Name, isVisible, isVeryHidden));
                    }
                }
            }
            catch { }
            return list;
        }

        public void SetWorksheetVisibility(string sheetName, bool visible, bool veryHidden = false)
        {
            if (IsWorkbookProtected)
            {
                ShowMessage("The active workbook structure is protected. Please unprotect the workbook before changing worksheet visibility.", "Workbook Protected", isError: true);
                return;
            }

            try
            {
                if (_app.ActiveWorkbook != null)
                {
                    var ws = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[sheetName];
                    if (ws != null)
                    {
                        if (veryHidden) ws.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                        else if (visible) ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        else ws.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error updating visibility for sheet '{sheetName}': {ex.Message}", "Worksheet Error", isError: true);
            }
        }

        public void RenameWorksheets(IDictionary<string, string> renameMap)
        {
            if (renameMap == null) return;

            if (IsWorkbookProtected)
            {
                ShowMessage("The active workbook structure is protected. Please unprotect the workbook before renaming worksheets.", "Workbook Protected", isError: true);
                return;
            }

            using (new ExcelStateScope(_app))
            {
                // First pass: rename to temporary names to avoid transient collisions
                var tempNames = new Dictionary<string, string>();
                foreach (var pair in renameMap)
                {
                    if (!string.Equals(pair.Key, pair.Value, StringComparison.OrdinalIgnoreCase))
                    {
                        string tempName = $"__EPS_TEMP_{Guid.NewGuid():N}".Substring(0, 28);
                        try
                        {
                            var ws = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[pair.Key];
                            ws.Name = tempName;
                            tempNames[tempName] = pair.Value;
                        }
                        catch { }
                    }
                }

                // Second pass: apply final desired names
                foreach (var pair in tempNames)
                {
                    try
                    {
                        var ws = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[pair.Key];
                        ws.Name = pair.Value;
                    }
                    catch { }
                }
            }
        }

        public bool CreateWorksheetWithData(string sheetName, object?[,] data, bool activate = true)
        {
            if (IsWorkbookProtected)
            {
                ShowMessage("The active workbook structure is protected. Cannot add new worksheets.", "Workbook Protected", isError: true);
                return false;
            }

            try
            {
                if (_app.ActiveWorkbook == null) return false;

                using (new ExcelStateScope(_app))
                {
                    // Uniquify sheet name against existing worksheets
                    var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (Excel.Worksheet w in _app.ActiveWorkbook.Worksheets)
                    {
                        existing.Add(w.Name);
                    }

                    string validName = sheetName;
                    if (existing.Contains(validName) || validName.Length > 31)
                    {
                        string baseName = validName.Length > 25 ? validName.Substring(0, 25) : validName;
                        int counter = 1;
                        do
                        {
                            validName = $"{baseName}_{counter++}";
                        } while (existing.Contains(validName));
                    }

                    Excel.Worksheet newSheet = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets.Add();
                    newSheet.Name = validName;

                    if (data != null)
                    {
                        int rows = data.GetLength(0);
                        int cols = data.GetLength(1);
                        if (rows > 0 && cols > 0)
                        {
                            Excel.Range topLeft = newSheet.Cells[1, 1];
                            Excel.Range bottomRight = newSheet.Cells[rows, cols];
                            Excel.Range target = newSheet.Range[topLeft, bottomRight];
                            target.Value2 = data;
                        }
                    }

                    if (activate)
                    {
                        newSheet.Activate();
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error creating worksheet '{sheetName}': {ex.Message}", "Worksheet Error", isError: true);
                return false;
            }
        }

        public IReadOnlyList<string> GetOpenWorkbookNames()
        {
            var list = new List<string>();
            try
            {
                foreach (Excel.Workbook wb in _app.Workbooks)
                {
                    list.Add(wb.Name);
                }
            }
            catch { }
            return list;
        }

        public IReadOnlyList<string> GetWorkbookSheetNames(string workbookName)
        {
            var list = new List<string>();
            try
            {
                Excel.Workbook? wb = null;
                foreach (Excel.Workbook b in _app.Workbooks)
                {
                    if (string.Equals(b.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                    {
                        wb = b;
                        break;
                    }
                }

                if (wb != null)
                {
                    foreach (Excel.Worksheet ws in wb.Worksheets)
                    {
                        list.Add(ws.Name);
                    }
                }
            }
            catch { }
            return list;
        }

        public object?[,]? GetWorkbookRangeValues(string workbookName, string sheetName, string? rangeAddress = null)
        {
            try
            {
                Excel.Workbook? wb = null;
                foreach (Excel.Workbook b in _app.Workbooks)
                {
                    if (string.Equals(b.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                    {
                        wb = b;
                        break;
                    }
                }

                if (wb == null) return null;

                Excel.Worksheet? ws = null;
                foreach (Excel.Worksheet s in wb.Worksheets)
                {
                    if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        ws = s;
                        break;
                    }
                }

                if (ws == null) return null;

                Excel.Range target = !string.IsNullOrWhiteSpace(rangeAddress) 
                    ? ws.Range[rangeAddress] 
                    : ws.UsedRange;

                if (target == null) return null;

                object raw = target.Value2;
                if (raw is object[,] matrix)
                {
                    return (object?[,])matrix;
                }

                object?[,] single = new object?[1, 1];
                single[0, 0] = raw;
                return single;
            }
            catch (Exception ex)
            {
                ShowMessage($"Unable to read range from '{workbookName} - {sheetName}': {ex.Message}", "Read Error", isError: true);
                return null;
            }
        }

        public IReadOnlyList<WorkbookSummaryInfo> GetOpenWorkbookDetails()
        {
            var list = new List<WorkbookSummaryInfo>();
            try
            {
                Excel.Workbook? activeWb = null;
                try { activeWb = _app.ActiveWorkbook; } catch { }

                foreach (Excel.Workbook wb in _app.Workbooks)
                {
                    try
                    {
                        string name = wb.Name;
                        string path = wb.FullName;
                        bool isActive = activeWb != null && string.Equals(activeWb.Name, name, StringComparison.OrdinalIgnoreCase);
                        bool isSaved = wb.Saved;
                        int sheetCount = wb.Sheets.Count;
                        bool isReadOnly = wb.ReadOnly;
                        list.Add(new WorkbookSummaryInfo(name, path, isActive, isSaved, sheetCount, isReadOnly));
                    }
                    catch
                    {
                        list.Add(new WorkbookSummaryInfo(wb.Name, null, false, true, 0));
                    }
                }
            }
            catch { }
            return list;
        }

        public IReadOnlyList<WorksheetSummaryInfo> GetWorksheetDetails(string? workbookName = null)
        {
            var list = new List<WorksheetSummaryInfo>();
            try
            {
                Excel.Workbook? wb = null;
                if (!string.IsNullOrWhiteSpace(workbookName))
                {
                    foreach (Excel.Workbook b in _app.Workbooks)
                    {
                        if (string.Equals(b.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            wb = b;
                            break;
                        }
                    }
                }
                else
                {
                    wb = _app.ActiveWorkbook;
                }

                if (wb == null) return list;

                Excel.Worksheet? activeWs = null;
                try { activeWs = wb.ActiveSheet as Excel.Worksheet; } catch { }

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    try
                    {
                        string name = ws.Name;
                        int index = ws.Index;
                        bool isActive = activeWs != null && string.Equals(activeWs.Name, name, StringComparison.OrdinalIgnoreCase);
                        SheetVisibilityState vis = ws.Visible switch
                        {
                            Excel.XlSheetVisibility.xlSheetVisible => SheetVisibilityState.Visible,
                            Excel.XlSheetVisibility.xlSheetHidden => SheetVisibilityState.Hidden,
                            Excel.XlSheetVisibility.xlSheetVeryHidden => SheetVisibilityState.VeryHidden,
                            _ => SheetVisibilityState.Visible
                        };

                        int? tabColor = null;
                        try
                        {
                            if (ws.Tab != null)
                            {
                                object? c = ws.Tab.Color;
                                if (c != null) tabColor = Convert.ToInt32(c);
                            }
                        }
                        catch { }

                        bool isProtected = false;
                        try { isProtected = ws.ProtectContents; } catch { }

                        list.Add(new WorksheetSummaryInfo(name, wb.Name, index, isActive, vis, tabColor, isProtected));
                    }
                    catch { }
                }
            }
            catch { }
            return list;
        }

        public bool ActivateWorkbook(string workbookName)
        {
            try
            {
                foreach (Excel.Workbook wb in _app.Workbooks)
                {
                    if (string.Equals(wb.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                    {
                        wb.Activate();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not activate workbook '{workbookName}': {ex.Message}", "Activate Workbook", isError: true);
            }
            return false;
        }

        public bool CloseWorkbook(string workbookName, bool saveChanges = false)
        {
            try
            {
                foreach (Excel.Workbook wb in _app.Workbooks)
                {
                    if (string.Equals(wb.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                    {
                        wb.Close(saveChanges);
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not close workbook '{workbookName}': {ex.Message}", "Close Workbook", isError: true);
            }
            return false;
        }

        public bool SaveWorkbook(string workbookName)
        {
            try
            {
                foreach (Excel.Workbook wb in _app.Workbooks)
                {
                    if (string.Equals(wb.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                    {
                        wb.Save();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not save workbook '{workbookName}': {ex.Message}", "Save Workbook", isError: true);
            }
            return false;
        }

        public void SaveAllWorkbooks()
        {
            try
            {
                foreach (Excel.Workbook wb in _app.Workbooks)
                {
                    try
                    {
                        if (!wb.Saved) wb.Save();
                    }
                    catch { }
                }
            }
            catch { }
        }

        public bool CreateNewWorkbook()
        {
            try
            {
                var wb = _app.Workbooks.Add();
                wb.Activate();
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not create new workbook: {ex.Message}", "New Workbook", isError: true);
                return false;
            }
        }

        public bool ActivateWorksheet(string sheetName, string? workbookName = null)
        {
            try
            {
                Excel.Workbook? wb = null;
                if (!string.IsNullOrWhiteSpace(workbookName))
                {
                    foreach (Excel.Workbook b in _app.Workbooks)
                    {
                        if (string.Equals(b.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            wb = b;
                            break;
                        }
                    }
                }
                else
                {
                    wb = _app.ActiveWorkbook;
                }

                if (wb == null) return false;

                try
                {
                    if (wb.ActiveSheet is Excel.Worksheet curWs)
                    {
                        if (!string.Equals(curWs.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            _previousActiveSheetName = curWs.Name;
                        }
                    }
                }
                catch { }

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        }
                        ws.Activate();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not switch to worksheet '{sheetName}': {ex.Message}", "Activate Worksheet", isError: true);
            }
            return false;
        }

        public bool MoveWorksheet(string sheetName, SheetMoveDirection direction, string? workbookName = null)
        {
            try
            {
                Excel.Workbook? wb = null;
                if (!string.IsNullOrWhiteSpace(workbookName))
                {
                    foreach (Excel.Workbook b in _app.Workbooks)
                    {
                        if (string.Equals(b.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            wb = b;
                            break;
                        }
                    }
                }
                else
                {
                    wb = _app.ActiveWorkbook;
                }

                if (wb == null) return false;

                Excel.Worksheet? ws = null;
                foreach (Excel.Worksheet s in wb.Worksheets)
                {
                    if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        ws = s;
                        break;
                    }
                }

                if (ws == null) return false;

                int curIndex = ws.Index;
                int totalSheets = wb.Sheets.Count;

                switch (direction)
                {
                    case SheetMoveDirection.Top:
                        if (curIndex > 1)
                        {
                            ws.Move(Before: wb.Sheets[1]);
                        }
                        break;

                    case SheetMoveDirection.End:
                        if (curIndex < totalSheets)
                        {
                            ws.Move(After: wb.Sheets[totalSheets]);
                        }
                        break;

                    case SheetMoveDirection.Up:
                        if (curIndex > 1)
                        {
                            ws.Move(Before: wb.Sheets[curIndex - 1]);
                        }
                        break;

                    case SheetMoveDirection.Down:
                        if (curIndex < totalSheets)
                        {
                            ws.Move(After: wb.Sheets[curIndex + 1]);
                        }
                        break;
                }
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not move worksheet '{sheetName}': {ex.Message}", "Move Worksheet", isError: true);
                return false;
            }
        }

        public bool ToggleLastActiveWorksheet(string? workbookName = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_previousActiveSheetName)) return false;
                string target = _previousActiveSheetName!;
                return ActivateWorksheet(target, workbookName);
            }
            catch
            {
                return false;
            }
        }

        public bool AddWorksheet(string? desiredName = null, string? workbookName = null)
        {
            try
            {
                Excel.Workbook? wb = null;
                if (!string.IsNullOrWhiteSpace(workbookName))
                {
                    foreach (Excel.Workbook b in _app.Workbooks)
                    {
                        if (string.Equals(b.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            wb = b;
                            break;
                        }
                    }
                }
                else
                {
                    wb = _app.ActiveWorkbook;
                }

                if (wb == null) return false;

                var ws = (Excel.Worksheet)wb.Worksheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                if (!string.IsNullOrWhiteSpace(desiredName))
                {
                    try { ws.Name = desiredName; } catch { }
                }
                ws.Activate();
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not add worksheet: {ex.Message}", "Add Worksheet", isError: true);
                return false;
            }
        }

        public bool DeleteWorksheet(string sheetName, string? workbookName = null)
        {
            try
            {
                Excel.Workbook? wb = null;
                if (!string.IsNullOrWhiteSpace(workbookName))
                {
                    foreach (Excel.Workbook b in _app.Workbooks)
                    {
                        if (string.Equals(b.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            wb = b;
                            break;
                        }
                    }
                }
                else
                {
                    wb = _app.ActiveWorkbook;
                }

                if (wb == null) return false;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        bool origAlerts = _app.DisplayAlerts;
                        try
                        {
                            _app.DisplayAlerts = false;
                            ws.Delete();
                        }
                        finally
                        {
                            _app.DisplayAlerts = origAlerts;
                        }
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not delete worksheet '{sheetName}': {ex.Message}", "Delete Worksheet", isError: true);
            }
            return false;
        }

        public bool RenameWorksheet(string oldName, string newName, string? workbookName = null)
        {
            try
            {
                Excel.Workbook? wb = null;
                if (!string.IsNullOrWhiteSpace(workbookName))
                {
                    foreach (Excel.Workbook b in _app.Workbooks)
                    {
                        if (string.Equals(b.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            wb = b;
                            break;
                        }
                    }
                }
                else
                {
                    wb = _app.ActiveWorkbook;
                }

                if (wb == null) return false;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, oldName, StringComparison.OrdinalIgnoreCase))
                    {
                        ws.Name = newName;
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not rename worksheet '{oldName}' to '{newName}': {ex.Message}", "Rename Worksheet", isError: true);
            }
            return false;
        }

        public void ShowNavigationPane()
        {
            try
            {
                if (NavigationPaneToggleAction != null)
                {
                    NavigationPaneToggleAction.Invoke();
                }
                else
                {
                    var win = new ExcelProductivitySuite.UI.Dialogs.WorkbookSheetNavigatorWindow(this);
                    win.Show();
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not display Navigation Pane: {ex.Message}", "Navigation Pane Error", isError: true);
            }
        }

        public void EnableReadingView(ReadingViewSettings? settings = null)
        {
            if (settings != null)
            {
                _readingViewSettings = settings.Clone();
            }
            _readingViewSettings.IsEnabled = true;
            _isReadingViewActive = true;
            ReadingLayoutEngine.SaveSettings(_readingViewSettings);
            UpdateReadingViewHighlight();
        }

        public void DisableReadingView()
        {
            _isReadingViewActive = false;
            _readingViewSettings.IsEnabled = false;
            ReadingLayoutEngine.SaveSettings(_readingViewSettings);
            try
            {
                if (HasActiveWorksheet && _app.ActiveSheet is Excel.Worksheet ws)
                {
                    RemoveReadingViewShapes(ws);
                }
            }
            catch { }
        }

        public void ToggleReadingView()
        {
            if (_isReadingViewActive)
            {
                DisableReadingView();
            }
            else
            {
                EnableReadingView();
            }
        }

        public void EnableGridFocus(ReadingViewSettings? settings = null) => EnableReadingView(settings);
        public void DisableGridFocus() => DisableReadingView();
        public void ToggleGridFocus() => ToggleReadingView();

        private void RemoveReadingViewShapes(Excel.Worksheet ws)
        {
            try
            {
                string[] names = { ReadingLayoutEngine.ShapeNameRow, ReadingLayoutEngine.ShapeNameCol, ReadingLayoutEngine.ShapeNameCell };
                foreach (var name in names)
                {
                    try
                    {
                        var shape = ws.Shapes.Item(name);
                        if (shape != null)
                        {
                            shape.Delete();
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        public void UpdateReadingViewHighlight()
        {
            if (!_isReadingViewActive || !HasActiveWorksheet) return;

            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                var cell = _app.ActiveCell;
                if (cell == null) return;

                double cellLeft = Convert.ToDouble(cell.Left);
                double cellTop = Convert.ToDouble(cell.Top);
                double cellWidth = Convert.ToDouble(cell.Width);
                double cellHeight = Convert.ToDouble(cell.Height);

                double maxLeft = 0;
                double maxTop = 0;
                double maxWidth = 3000;
                double maxHeight = 30000;

                if (_readingViewSettings.Scope == ReadingHighlightScope.UsedRangeOnly)
                {
                    try
                    {
                        var ur = ws.UsedRange;
                        if (ur != null)
                        {
                            maxLeft = Convert.ToDouble(ur.Left);
                            maxTop = Convert.ToDouble(ur.Top);
                            maxWidth = Convert.ToDouble(ur.Width);
                            maxHeight = Convert.ToDouble(ur.Height);
                        }
                    }
                    catch { }
                }

                var geo = ReadingLayoutEngine.CalculateGeometry(
                    cellLeft, cellTop, cellWidth, cellHeight,
                    maxLeft, maxTop, maxWidth, maxHeight,
                    _readingViewSettings);

                // Remove existing shapes
                RemoveReadingViewShapes(ws);

                int colorRgb = _readingViewSettings.ColorRgb;
                float transparency = (float)Math.Max(0.0, Math.Min(1.0, _readingViewSettings.Transparency));

                // 1. Row Shape
                if (geo.HasRow)
                {
                    try
                    {
                        var rowShape = ws.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            (float)geo.RowLeft, (float)geo.RowTop, (float)geo.RowWidth, (float)geo.RowHeight);
                        rowShape.Name = ReadingLayoutEngine.ShapeNameRow;
                        rowShape.Placement = Excel.XlPlacement.xlFreeFloating;
                        rowShape.Fill.Solid();
                        rowShape.Fill.ForeColor.RGB = colorRgb;
                        rowShape.Fill.Transparency = transparency;
                        if (_readingViewSettings.BorderEnabled)
                        {
                            rowShape.Line.Visible = Office.MsoTriState.msoTrue;
                            rowShape.Line.ForeColor.RGB = _readingViewSettings.BorderColorRgb;
                            rowShape.Line.Weight = 1.0f;
                        }
                        else
                        {
                            rowShape.Line.Visible = Office.MsoTriState.msoFalse;
                        }
                    }
                    catch { }
                }

                // 2. Column Shape
                if (geo.HasColumn)
                {
                    try
                    {
                        var colShape = ws.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            (float)geo.ColLeft, (float)geo.ColTop, (float)geo.ColWidth, (float)geo.ColHeight);
                        colShape.Name = ReadingLayoutEngine.ShapeNameCol;
                        colShape.Placement = Excel.XlPlacement.xlFreeFloating;
                        colShape.Fill.Solid();
                        colShape.Fill.ForeColor.RGB = colorRgb;
                        colShape.Fill.Transparency = transparency;
                        if (_readingViewSettings.BorderEnabled)
                        {
                            colShape.Line.Visible = Office.MsoTriState.msoTrue;
                            colShape.Line.ForeColor.RGB = _readingViewSettings.BorderColorRgb;
                            colShape.Line.Weight = 1.0f;
                        }
                        else
                        {
                            colShape.Line.Visible = Office.MsoTriState.msoFalse;
                        }
                    }
                    catch { }
                }

                // 3. Active Cell Focus Border
                if (geo.HasCell && _readingViewSettings.BorderEnabled)
                {
                    try
                    {
                        var cellShape = ws.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            (float)geo.CellLeft, (float)geo.CellTop, (float)geo.CellWidth, (float)geo.CellHeight);
                        cellShape.Name = ReadingLayoutEngine.ShapeNameCell;
                        cellShape.Placement = Excel.XlPlacement.xlFreeFloating;
                        cellShape.Fill.Visible = Office.MsoTriState.msoFalse;
                        cellShape.Line.Visible = Office.MsoTriState.msoTrue;
                        cellShape.Line.ForeColor.RGB = _readingViewSettings.BorderColorRgb;
                        cellShape.Line.Weight = 2.0f;
                    }
                    catch { }
                }
            }
            catch { }
        }

        public object?[,]? GetActiveWorksheetRangeValues(string sheetName, string? rangeAddress = null)
        {
            try
            {
                if (_app.ActiveWorkbook == null) return null;
                Excel.Worksheet? ws = null;
                foreach (Excel.Worksheet s in _app.ActiveWorkbook.Worksheets)
                {
                    if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        ws = s;
                        break;
                    }
                }

                if (ws == null) return null;

                Excel.Range target = !string.IsNullOrWhiteSpace(rangeAddress)
                    ? ws.Range[rangeAddress]
                    : ws.UsedRange;

                if (target == null) return null;

                object raw = target.Value2;
                if (raw is object[,] matrix)
                {
                    return (object?[,])matrix;
                }

                object?[,] single = new object?[1, 1];
                single[0, 0] = raw;
                return single;
            }
            catch (Exception ex)
            {
                ShowMessage($"Unable to read range from worksheet '{sheetName}': {ex.Message}", "Read Error", isError: true);
                return null;
            }
        }

        public void ReorderWorksheets(IList<string> orderedSheetNames)
        {
            if (orderedSheetNames == null || orderedSheetNames.Count == 0) return;
            if (IsWorkbookProtected)
            {
                ShowMessage("The active workbook structure is protected. Cannot reorder worksheets.", "Workbook Protected", isError: true);
                return;
            }

            try
            {
                if (_app.ActiveWorkbook == null) return;
                using (new ExcelStateScope(_app))
                {
                    for (int i = 0; i < orderedSheetNames.Count; i++)
                    {
                        try
                        {
                            var ws = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[orderedSheetNames[i]];
                            if (ws != null)
                            {
                                ws.Move(After: _app.ActiveWorkbook.Worksheets[_app.ActiveWorkbook.Worksheets.Count]);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error reordering worksheets: {ex.Message}", "Worksheet Error", isError: true);
            }
        }

        public IReadOnlyList<string> GetBlankWorksheetNames()
        {
            var blankSheets = new List<string>();
            try
            {
                if (_app.ActiveWorkbook == null) return blankSheets;
                foreach (Excel.Worksheet ws in _app.ActiveWorkbook.Worksheets)
                {
                    bool isBlank = true;
                    try
                    {
                        Excel.Range used = ws.UsedRange;
                        if (used != null)
                        {
                            object val = used.Value2;
                            if (val is object[,] matrix)
                            {
                                int rStart = matrix.GetLowerBound(0);
                                int rEnd = matrix.GetUpperBound(0);
                                int cStart = matrix.GetLowerBound(1);
                                int cEnd = matrix.GetUpperBound(1);

                                for (int r = rStart; r <= rEnd && isBlank; r++)
                                {
                                    for (int c = cStart; c <= cEnd; c++)
                                    {
                                        var cellVal = matrix[r, c];
                                        if (cellVal != null && !string.IsNullOrWhiteSpace(cellVal.ToString()))
                                        {
                                            isBlank = false;
                                            break;
                                        }
                                    }
                                }
                            }
                            else if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                            {
                                isBlank = false;
                            }
                        }

                        if (isBlank && ws.Shapes != null && ws.Shapes.Count > 0)
                        {
                            isBlank = false;
                        }
                    }
                    catch
                    {
                        isBlank = false;
                    }

                    if (isBlank)
                    {
                        blankSheets.Add(ws.Name);
                    }
                }
            }
            catch { }
            return blankSheets;
        }

        public int DeleteWorksheets(IEnumerable<string> sheetNames)
        {
            if (sheetNames == null) return 0;
            if (IsWorkbookProtected)
            {
                ShowMessage("The active workbook structure is protected. Cannot delete worksheets.", "Workbook Protected", isError: true);
                return 0;
            }

            int deletedCount = 0;
            try
            {
                if (_app.ActiveWorkbook == null) return 0;

                var toDelete = new HashSet<string>(sheetNames, StringComparer.OrdinalIgnoreCase);
                int totalSheets = _app.ActiveWorkbook.Worksheets.Count;
                if (toDelete.Count >= totalSheets)
                {
                    ShowMessage("Cannot delete all worksheets in the workbook. At least one sheet must remain.", "Delete Canceled", isError: true);
                    return 0;
                }

                using (new ExcelStateScope(_app))
                {
                    _app.DisplayAlerts = false;
                    foreach (string sName in toDelete)
                    {
                        try
                        {
                            var ws = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[sName];
                            if (ws != null)
                            {
                                ws.Delete();
                                deletedCount++;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error deleting worksheets: {ex.Message}", "Delete Error", isError: true);
            }
            return deletedCount;
        }

        public bool CreateIndexWorksheet(string indexSheetName, IEnumerable<string> sheetNames, bool includeBackLinks = false)
        {
            if (sheetNames == null) return false;
            if (IsWorkbookProtected)
            {
                ShowMessage("The active workbook structure is protected. Cannot create Table of Contents worksheet.", "Workbook Protected", isError: true);
                return false;
            }

            try
            {
                if (_app.ActiveWorkbook == null) return false;

                using (new ExcelStateScope(_app))
                {
                    var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (Excel.Worksheet w in _app.ActiveWorkbook.Worksheets)
                    {
                        existing.Add(w.Name);
                    }

                    string validName = indexSheetName;
                    if (existing.Contains(validName) || validName.Length > 31)
                    {
                        string baseName = validName.Length > 25 ? validName.Substring(0, 25) : validName;
                        int counter = 1;
                        do
                        {
                            validName = $"{baseName}_{counter++}";
                        } while (existing.Contains(validName));
                    }

                    var firstSheet = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[1];
                    Excel.Worksheet indexSheet = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets.Add(Before: firstSheet);
                    indexSheet.Name = validName;

                    indexSheet.Cells[1, 1].Value2 = "#";
                    indexSheet.Cells[1, 2].Value2 = "Worksheet Name";
                    indexSheet.Cells[1, 3].Value2 = "Status";
                    indexSheet.Cells[1, 4].Value2 = "Quick Link";

                    Excel.Range headerRange = indexSheet.Range["A1", "D1"];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = 0xE8F5E9;

                    int rowIdx = 2;
                    int seq = 1;
                    foreach (string sName in sheetNames)
                    {
                        if (string.Equals(sName, validName, StringComparison.OrdinalIgnoreCase)) continue;

                        try
                        {
                            var targetWs = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[sName];
                            string status = targetWs.Visible == Excel.XlSheetVisibility.xlSheetVisible ? "Visible" : "Hidden";
                            string safeTargetName = sName.Replace("'", "''");

                            indexSheet.Cells[rowIdx, 1].Value2 = seq++;
                            indexSheet.Cells[rowIdx, 2].Value2 = sName;
                            indexSheet.Cells[rowIdx, 3].Value2 = status;

                            Excel.Range linkCell = indexSheet.Cells[rowIdx, 4];
                            linkCell.Formula = $"=HYPERLINK(\"#'{safeTargetName}'!A1\", \"Go to {safeTargetName}\")";

                            if (includeBackLinks && targetWs.Visible == Excel.XlSheetVisibility.xlSheetVisible && !targetWs.ProtectContents)
                            {
                                try
                                {
                                    Excel.Range backCell = targetWs.Range["Z1"];
                                    if (backCell.Value2 == null)
                                    {
                                        backCell.Formula = $"=HYPERLINK(\"#'{validName}'!A1\", \"[Back to Index]\")";
                                    }
                                }
                                catch { }
                            }

                            rowIdx++;
                        }
                        catch { }
                    }

                    indexSheet.Columns.AutoFit();
                    indexSheet.Activate();
                    return true;
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error creating Table of Contents: {ex.Message}", "TOC Error", isError: true);
                return false;
            }
        }

        public int SplitWorksheetsToFiles(IEnumerable<string> sheetNames, string targetDirectory, SplitFileFormat format, string namingPattern)
        {
            if (sheetNames == null || string.IsNullOrWhiteSpace(targetDirectory)) return 0;
            if (!System.IO.Directory.Exists(targetDirectory))
            {
                try { System.IO.Directory.CreateDirectory(targetDirectory); }
                catch (Exception ex)
                {
                    ShowMessage($"Cannot create output directory: {ex.Message}", "Directory Error", isError: true);
                    return 0;
                }
            }

            int exportedCount = 0;
            try
            {
                if (_app.ActiveWorkbook == null) return 0;
                string wbName = System.IO.Path.GetFileNameWithoutExtension(_app.ActiveWorkbook.Name);

                using (new ExcelStateScope(_app))
                {
                    _app.DisplayAlerts = false;
                    foreach (string sName in sheetNames)
                    {
                        try
                        {
                            var ws = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[sName];
                            if (ws == null) continue;

                            string sanitizedSheetName = string.Join("_", sName.Split(System.IO.Path.GetInvalidFileNameChars()));
                            string fileBaseName = namingPattern
                                .Replace("[SheetName]", sanitizedSheetName)
                                .Replace("[WorkbookName]", wbName);

                            if (string.IsNullOrWhiteSpace(fileBaseName)) fileBaseName = sanitizedSheetName;

                            switch (format)
                            {
                                case SplitFileFormat.ExcelWorkbook:
                                    string xlsxPath = System.IO.Path.Combine(targetDirectory, $"{fileBaseName}.xlsx");
                                    ws.Copy();
                                    var newWb = _app.ActiveWorkbook;
                                    newWb.SaveAs(xlsxPath, Excel.XlFileFormat.xlOpenXMLWorkbook);
                                    newWb.Close(false);
                                    exportedCount++;
                                    break;

                                case SplitFileFormat.Pdf:
                                    string pdfPath = System.IO.Path.Combine(targetDirectory, $"{fileBaseName}.pdf");
                                    ws.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfPath);
                                    exportedCount++;
                                    break;

                                case SplitFileFormat.Csv:
                                    string csvPath = System.IO.Path.Combine(targetDirectory, $"{fileBaseName}.csv");
                                    ws.Copy();
                                    var csvWb = _app.ActiveWorkbook;
                                    csvWb.SaveAs(csvPath, Excel.XlFileFormat.xlCSV);
                                    csvWb.Close(false);
                                    exportedCount++;
                                    break;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error splitting worksheets: {ex.Message}", "Split Error", isError: true);
            }

            return exportedCount;
        }

        public int ImportSheetsFromWorkbook(string filePath, bool firstSheetOnly, bool prefixWithFileName)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !System.IO.File.Exists(filePath))
            {
                ShowMessage($"File does not exist: {filePath}", "File Not Found", isError: true);
                return 0;
            }

            if (IsWorkbookProtected)
            {
                ShowMessage("The active workbook structure is protected. Cannot import worksheets.", "Workbook Protected", isError: true);
                return 0;
            }

            int importedCount = 0;
            Excel.Workbook? srcWb = null;
            try
            {
                if (_app.ActiveWorkbook == null) return 0;
                Excel.Workbook targetWb = _app.ActiveWorkbook;
                string filePrefix = System.IO.Path.GetFileNameWithoutExtension(filePath);

                using (new ExcelStateScope(_app))
                {
                    _app.DisplayAlerts = false;
                    srcWb = _app.Workbooks.Open(filePath, ReadOnly: true);
                    if (srcWb != null)
                    {
                        int sheetCount = srcWb.Worksheets.Count;
                        int limit = firstSheetOnly ? 1 : sheetCount;

                        for (int i = 1; i <= limit; i++)
                        {
                            try
                            {
                                var srcSheet = (Excel.Worksheet)srcWb.Worksheets[i];
                                string originalName = srcSheet.Name;
                                srcSheet.Copy(After: targetWb.Worksheets[targetWb.Worksheets.Count]);

                                var copiedSheet = (Excel.Worksheet)targetWb.Worksheets[targetWb.Worksheets.Count];
                                if (prefixWithFileName)
                                {
                                    string candidate = $"{filePrefix}_{originalName}";
                                    if (candidate.Length > 31) candidate = candidate.Substring(0, 31);
                                    try { copiedSheet.Name = candidate; } catch { }
                                }

                                importedCount++;
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error importing sheets from '{System.IO.Path.GetFileName(filePath)}': {ex.Message}", "Import Error", isError: true);
            }
            finally
            {
                if (srcWb != null)
                {
                    try { srcWb.Close(false); } catch { }
                    try { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(srcWb); } catch { }
                    srcWb = null;
                }
            }

            return importedCount;
        }

        public IReadOnlyList<CellColorInfo> GetSelectedCellColors(int maxCells = 5000)
        {
            var list = new List<CellColorInfo>();
            try
            {
                if (!(_app.Selection is Excel.Range range)) return list;
                int count = 0;

                foreach (Excel.Range area in range.Areas)
                {
                    int rows = area.Rows.Count;
                    int cols = area.Columns.Count;

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            try
                            {
                                Excel.Range cell = area.Cells[r, c];
                                int sheetRow = cell.Row;
                                int sheetCol = cell.Column;
                                object val = cell.Value2;

                                int fillColor = 0xFFFFFF;
                                if (cell.Interior != null && cell.Interior.Color is double fcd)
                                {
                                    fillColor = (int)fcd;
                                }

                                int fontColor = 0x000000;
                                if (cell.Font != null && cell.Font.Color is double fnd)
                                {
                                    fontColor = (int)fnd;
                                }

                                list.Add(new CellColorInfo(sheetRow, sheetCol, fillColor, fontColor, val));
                                count++;
                                if (count >= maxCells) return list;
                            }
                            catch { }
                        }
                    }
                }
            }
            catch { }
            return list;
        }

        public bool ApplyPreventDuplicateValidation(string? customErrorMessage = null)
        {
            try
            {
                if (!(_app.Selection is Excel.Range range)) return false;
                if (IsCurrentSheetProtected)
                {
                    ShowMessage("The active worksheet is protected. Cannot apply validation rules.", "Sheet Protected", isError: true);
                    return false;
                }

                using (new ExcelStateScope(_app))
                {
                    int startCol = range.Column;
                    string colLetter = GetColumnLetter(startCol);
                    int startRow = range.Row;

                    string formula = $"=COUNTIF({colLetter}:{colLetter}, {colLetter}{startRow})<=1";

                    Excel.Validation val = range.Validation;
                    val.Delete();
                    val.Add(
                        Excel.XlDVType.xlValidateCustom,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlBetween,
                        formula,
                        Type.Missing);

                    val.IgnoreBlank = true;
                    val.ErrorTitle = "Duplicate Entry Prohibited";
                    val.ErrorMessage = !string.IsNullOrWhiteSpace(customErrorMessage)
                        ? customErrorMessage
                        : "Duplicate entries are not permitted in this range.";
                    val.ShowError = true;

                    return true;
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error applying duplicate validation: {ex.Message}", "Validation Error", isError: true);
                return false;
            }
        }

        private static string GetColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = string.Empty;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return colLetter;
        }

        public object?[,]? GetSelectedFormulas()
        {
            try
            {
                if (!(_app.Selection is Excel.Range range)) return null;

                object raw = range.Formula;
                if (raw is object[,] matrix)
                {
                    return (object?[,])matrix;
                }

                object?[,] single = new object?[1, 1];
                single[0, 0] = raw;
                return single;
            }
            catch (Exception ex)
            {
                ShowMessage($"Error reading selection formulas: {ex.Message}", "Formula Error", isError: true);
                return null;
            }
        }

        public void SetSelectedFormulas(object?[,] formulas)
        {
            if (formulas == null) throw new ArgumentNullException(nameof(formulas));
            if (!(_app.Selection is Excel.Range range)) return;

            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before modifying formulas.", "Sheet Protected", isError: true);
                return;
            }

            if (IsMultiAreaSelection)
            {
                ShowMessage("This operation does not support non-contiguous multi-area selections.", "Multi-Area Selection", isError: true);
                return;
            }

            RecordBeforeRangeChange(range, formulas, isFormula: true, customDesc: "Formula Update");

            using (new ExcelStateScope(_app))
            {
                range.Formula = formulas;
            }
        }

        public object?[,]? GetRangeFormulas(int startRow, int startCol, int rowCount, int colCount)
        {
            try
            {
                if (!(_app.ActiveSheet is Excel.Worksheet sheet)) return null;

                Excel.Range topLeft = sheet.Cells[startRow, startCol];
                Excel.Range bottomRight = sheet.Cells[startRow + rowCount - 1, startCol + colCount - 1];
                Excel.Range targetRange = sheet.Range[topLeft, bottomRight];

                object raw = targetRange.Formula;
                if (raw is object[,] matrix)
                {
                    return (object?[,])matrix;
                }

                object?[,] single = new object?[1, 1];
                single[0, 0] = raw;
                return single;
            }
            catch (Exception ex)
            {
                ShowMessage($"Error reading range formulas: {ex.Message}", "Formula Error", isError: true);
                return null;
            }
        }

        public object?[,]? GetRangeFormulasByAddress(string address)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(address)) return null;
                Excel.Range targetRange = _app.Range[address];

                object raw = targetRange.Formula;
                if (raw is object[,] matrix)
                {
                    return (object?[,])matrix;
                }

                object?[,] single = new object?[1, 1];
                single[0, 0] = raw;
                return single;
            }
            catch (Exception ex)
            {
                ShowMessage($"Error reading formulas by address '{address}': {ex.Message}", "Formula Error", isError: true);
                return null;
            }
        }

        public void SetRangeFormulas(int startRow, int startCol, object?[,] formulas)
        {
            if (formulas == null) throw new ArgumentNullException(nameof(formulas));
            if (!(_app.ActiveSheet is Excel.Worksheet sheet)) return;

            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before modifying formulas.", "Sheet Protected", isError: true);
                return;
            }

            int rows = formulas.GetLength(0);
            int cols = formulas.GetLength(1);
            if (rows <= 0 || cols <= 0) return;

            using (new ExcelStateScope(_app))
            {
                Excel.Range topLeft = sheet.Cells[startRow, startCol];
                Excel.Range bottomRight = sheet.Cells[startRow + rows - 1, startCol + cols - 1];
                Excel.Range targetRange = sheet.Range[topLeft, bottomRight];

                RecordBeforeRangeChange(targetRange, formulas, isFormula: true, customDesc: "Formula Update");

                targetRange.Formula = formulas;
            }
        }

        public void SetRangeFormulasByAddress(string address, object?[,] formulas)
        {
            if (string.IsNullOrWhiteSpace(address)) return;
            if (formulas == null) throw new ArgumentNullException(nameof(formulas));

            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before modifying formulas.", "Sheet Protected", isError: true);
                return;
            }

            using (new ExcelStateScope(_app))
            {
                Excel.Range targetRange = _app.Range[address];
                int rows = formulas.GetLength(0);
                int cols = formulas.GetLength(1);
                if (rows > 0 && cols > 0 && (targetRange.Rows.Count != rows || targetRange.Columns.Count != cols))
                {
                    targetRange = targetRange.Resize[rows, cols];
                }

                RecordBeforeRangeChange(targetRange, formulas, isFormula: true, customDesc: "Formula Update");

                targetRange.Formula = formulas;
            }
        }

        public void CopyRangeFormatsByAddress(string sourceAddress, string destinationAddress)
        {
            if (string.IsNullOrWhiteSpace(sourceAddress) || string.IsNullOrWhiteSpace(destinationAddress)) return;
            if (!HasActiveWorkbook) return;

            try
            {
                if (IsCurrentSheetProtected) return;

                using (new ExcelStateScope(_app))
                {
                    Excel.Range srcRange = _app.Range[sourceAddress];
                    Excel.Range dstRange = _app.Range[destinationAddress];

                    if (dstRange.Rows.Count == 1 && dstRange.Columns.Count == 1 &&
                        (srcRange.Rows.Count > 1 || srcRange.Columns.Count > 1))
                    {
                        dstRange = dstRange.Resize[srcRange.Rows.Count, srcRange.Columns.Count];
                    }

                    srcRange.Copy();
                    dstRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                    _app.CutCopyMode = (Excel.XlCutCopyMode)0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to copy range formats: {ex.Message}");
            }
        }

        public string? PromptForRangeAddress(string prompt, string title, string defaultAddress = "")
        {
            if (!HasActiveWorkbook) return null;

            try
            {
                _app.ScreenUpdating = true;
                _app.DisplayAlerts = true;

                string def = string.IsNullOrWhiteSpace(defaultAddress) ? (SelectionAddress ?? "$A$1") : defaultAddress;

                object result = _app.InputBox(
                    Prompt: prompt,
                    Title: title,
                    Default: def,
                    Type: 8);

                if (result is Excel.Range range)
                {
                    string activeSheetName = "";
                    try { activeSheetName = ((Excel.Worksheet)_app.ActiveSheet).Name; } catch { }

                    string rangeSheetName = "";
                    try { rangeSheetName = ((Excel.Worksheet)range.Parent).Name; } catch { }

                    string addr = range.get_Address(Type.Missing, Type.Missing, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                    if (!string.IsNullOrEmpty(rangeSheetName) && !string.Equals(rangeSheetName, activeSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return $"'{rangeSheetName}'!{addr}";
                    }

                    return addr;
                }
            }
            catch (COMException)
            {
                // User cancelled Excel InputBox dialog
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Range selection prompt error: {ex.Message}");
                return null;
            }

            return null;
        }

        public IReadOnlyList<FormulaErrorInfo> GetFormulaErrors(bool selectionOnly = false)
        {
            var results = new List<FormulaErrorInfo>();
            try
            {
                if (!(_app.ActiveSheet is Excel.Worksheet sheet)) return results;

                Excel.Range targetRange = selectionOnly && _app.Selection is Excel.Range sel ? sel : sheet.UsedRange;
                if (targetRange == null) return results;

                int startRow = targetRange.Row;
                int startCol = targetRange.Column;
                string sheetName = sheet.Name;

                object rawValues = targetRange.Value2;
                object rawFormulas = targetRange.Formula;

                object?[,]? valMatrix = rawValues as object[,];
                if (valMatrix == null && rawValues != null)
                {
                    valMatrix = new object?[1, 1] { { rawValues } };
                }

                object?[,]? formMatrix = rawFormulas as object[,];
                if (formMatrix == null && rawFormulas != null)
                {
                    formMatrix = new object?[1, 1] { { rawFormulas } };
                }

                if (valMatrix != null)
                {
                    results.AddRange(FormulaErrorLocator.LocateErrors(valMatrix, formMatrix, startRow, startCol, sheetName));
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error locating formula errors: {ex.Message}", "Formula Audit", isError: true);
            }

            return results;
        }

        public int CopyPageSetup(string sourceSheetName, IEnumerable<string> targetSheetNames, PageSetupOptions options)
        {
            if (string.IsNullOrWhiteSpace(sourceSheetName)) throw new ArgumentNullException(nameof(sourceSheetName));
            if (targetSheetNames == null) throw new ArgumentNullException(nameof(targetSheetNames));
            if (options == null) throw new ArgumentNullException(nameof(options));

            try
            {
                if (!(_app.ActiveWorkbook is Excel.Workbook wb)) return 0;

                Excel.Worksheet? srcWs = null;
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, sourceSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        srcWs = ws;
                        break;
                    }
                }

                if (srcWs == null)
                {
                    ShowMessage($"Source worksheet '{sourceSheetName}' not found.", "Page Setup Error", isError: true);
                    return 0;
                }

                Excel.PageSetup src = srcWs.PageSetup;
                int count = 0;

                using (new ExcelStateScope(_app))
                {
                    foreach (string tName in targetSheetNames)
                    {
                        if (string.Equals(tName, sourceSheetName, StringComparison.OrdinalIgnoreCase)) continue;

                        Excel.Worksheet? targetWs = null;
                        foreach (Excel.Worksheet ws in wb.Worksheets)
                        {
                            if (string.Equals(ws.Name, tName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetWs = ws;
                                break;
                            }
                        }

                        if (targetWs == null) continue;

                        try
                        {
                            Excel.PageSetup dst = targetWs.PageSetup;

                            if (options.CopyOrientationAndSize)
                            {
                                dst.Orientation = src.Orientation;
                                dst.PaperSize = src.PaperSize;
                            }

                            if (options.CopyMargins)
                            {
                                dst.TopMargin = src.TopMargin;
                                dst.BottomMargin = src.BottomMargin;
                                dst.LeftMargin = src.LeftMargin;
                                dst.RightMargin = src.RightMargin;
                                dst.HeaderMargin = src.HeaderMargin;
                                dst.FooterMargin = src.FooterMargin;
                            }

                            if (options.CopyPrintTitles)
                            {
                                dst.PrintTitleRows = src.PrintTitleRows;
                                dst.PrintTitleColumns = src.PrintTitleColumns;
                            }

                            if (options.CopyScaling)
                            {
                                dst.Zoom = src.Zoom;
                                dst.FitToPagesWide = src.FitToPagesWide;
                                dst.FitToPagesTall = src.FitToPagesTall;
                            }

                            if (options.CopyHeadersFooters)
                            {
                                dst.LeftHeader = src.LeftHeader;
                                dst.CenterHeader = src.CenterHeader;
                                dst.RightHeader = src.RightHeader;
                                dst.LeftFooter = src.LeftFooter;
                                dst.CenterFooter = src.CenterFooter;
                                dst.RightFooter = src.RightFooter;
                            }

                            if (options.CopySheetOptions)
                            {
                                dst.PrintGridlines = src.PrintGridlines;
                                dst.CenterHorizontally = src.CenterHorizontally;
                                dst.CenterVertically = src.CenterVertically;
                                dst.BlackAndWhite = src.BlackAndWhite;
                                dst.Draft = src.Draft;
                            }

                            count++;
                        }
                        catch
                        {
                            // Continue copying to next worksheet if one sheet is protected or restricted
                        }
                    }
                }

                return count;
            }
            catch (Exception ex)
            {
                ShowMessage($"Error copying page setup: {ex.Message}", "Page Setup Error", isError: true);
                return 0;
            }
        }

        public int SynchronizeFreezePanes(IEnumerable<string> targetSheetNames, FreezePanesSyncMode mode, int splitRow = 1, int splitCol = 1)
        {
            if (targetSheetNames == null) throw new ArgumentNullException(nameof(targetSheetNames));

            try
            {
                if (!(_app.ActiveWorkbook is Excel.Workbook wb)) return 0;
                string? originalActiveSheet = ActiveWorksheetName;
                int count = 0;

                using (new ExcelStateScope(_app))
                {
                    foreach (string tName in targetSheetNames)
                    {
                        Excel.Worksheet? targetWs = null;
                        foreach (Excel.Worksheet ws in wb.Worksheets)
                        {
                            if (string.Equals(ws.Name, tName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetWs = ws;
                                break;
                            }
                        }

                        if (targetWs == null) continue;

                        try
                        {
                            targetWs.Activate();
                            Excel.Window win = _app.ActiveWindow;
                            if (win == null) continue;

                            win.FreezePanes = false;

                            switch (mode)
                            {
                                case FreezePanesSyncMode.FreezeTopRow:
                                    win.SplitRow = 1;
                                    win.SplitColumn = 0;
                                    win.FreezePanes = true;
                                    break;
                                case FreezePanesSyncMode.FreezeFirstColumn:
                                    win.SplitRow = 0;
                                    win.SplitColumn = 1;
                                    win.FreezePanes = true;
                                    break;
                                case FreezePanesSyncMode.FreezeAtActiveCell:
                                    win.SplitRow = Math.Max(0, splitRow - 1);
                                    win.SplitColumn = Math.Max(0, splitCol - 1);
                                    win.FreezePanes = true;
                                    break;
                                case FreezePanesSyncMode.UnfreezeAll:
                                    win.FreezePanes = false;
                                    win.SplitRow = 0;
                                    win.SplitColumn = 0;
                                    break;
                            }

                            count++;
                        }
                        catch
                        {
                            // Skip sheets that cannot be activated or frozen
                        }
                    }

                    // Restore original sheet view
                    if (!string.IsNullOrEmpty(originalActiveSheet))
                    {
                        foreach (Excel.Worksheet ws in wb.Worksheets)
                        {
                            if (string.Equals(ws.Name, originalActiveSheet, StringComparison.OrdinalIgnoreCase))
                            {
                                ws.Activate();
                                break;
                            }
                        }
                    }
                }

                return count;
            }
            catch (Exception ex)
            {
                ShowMessage($"Error synchronizing freeze panes: {ex.Message}", "Freeze Panes Error", isError: true);
                return 0;
            }
        }

        private class PageSetupState
        {
            public Excel.XlPageOrientation Orientation { get; set; }
            public Excel.XlPaperSize PaperSize { get; set; }
            public object Zoom { get; set; } = 100;
            public object FitToPagesWide { get; set; } = false;
            public object FitToPagesTall { get; set; } = false;
        }

        private static PageSetupState CapturePageSetup(Excel.PageSetup ps)
        {
            return new PageSetupState
            {
                Orientation = ps.Orientation,
                PaperSize = ps.PaperSize,
                Zoom = ps.Zoom,
                FitToPagesWide = ps.FitToPagesWide,
                FitToPagesTall = ps.FitToPagesTall
            };
        }

        private static void RestorePageSetup(Excel.PageSetup ps, PageSetupState state)
        {
            try
            {
                ps.Orientation = state.Orientation;
                ps.PaperSize = state.PaperSize;
                ps.Zoom = state.Zoom;
                ps.FitToPagesWide = state.FitToPagesWide;
                ps.FitToPagesTall = state.FitToPagesTall;
            }
            catch { /* Ignore COM quirks during restore */ }
        }

        private static void ApplyPageSetupOverrides(Excel.PageSetup ps, PdfExportSettings settings)
        {
            try
            {
                if (settings.Orientation == PdfPageOrientation.Portrait)
                {
                    ps.Orientation = Excel.XlPageOrientation.xlPortrait;
                }
                else if (settings.Orientation == PdfPageOrientation.Landscape)
                {
                    ps.Orientation = Excel.XlPageOrientation.xlLandscape;
                }
            }
            catch { }

            try
            {
                if (settings.PaperSize != PdfPaperSize.Default)
                {
                    Excel.XlPaperSize? psEnum = settings.PaperSize switch
                    {
                        PdfPaperSize.A4 => Excel.XlPaperSize.xlPaperA4,
                        PdfPaperSize.Letter => Excel.XlPaperSize.xlPaperLetter,
                        PdfPaperSize.Legal => Excel.XlPaperSize.xlPaperLegal,
                        PdfPaperSize.A3 => Excel.XlPaperSize.xlPaperA3,
                        PdfPaperSize.Tabloid => Excel.XlPaperSize.xlPaper11x17,
                        PdfPaperSize.Executive => Excel.XlPaperSize.xlPaperExecutive,
                        _ => null
                    };
                    if (psEnum.HasValue)
                    {
                        ps.PaperSize = psEnum.Value;
                    }
                }
            }
            catch { }

            try
            {
                if (settings.PageFit != PdfPageFit.Default)
                {
                    ps.Zoom = false; // Zoom must be false for FitToPages to take effect in Excel
                    switch (settings.PageFit)
                    {
                        case PdfPageFit.FitSheetOnOnePage:
                            ps.FitToPagesWide = 1;
                            ps.FitToPagesTall = 1;
                            break;
                        case PdfPageFit.FitAllColumnsOnOnePage:
                            ps.FitToPagesWide = 1;
                            ps.FitToPagesTall = false;
                            break;
                        case PdfPageFit.FitAllRowsOnOnePage:
                            ps.FitToPagesWide = false;
                            ps.FitToPagesTall = 1;
                            break;
                    }
                }
            }
            catch { }
        }

        public bool ExportToPdf(string destinationPath, IEnumerable<string>? sheetNames = null, bool openAfterExport = false)
        {
            if (string.IsNullOrWhiteSpace(destinationPath)) throw new ArgumentNullException(nameof(destinationPath));

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

            try
            {
                if (!(_app.ActiveWorkbook is Excel.Workbook wb))
                {
                    ShowMessage("No active workbook is currently open.", "PDF Export Error", isError: true);
                    return PdfExportResult.Failed("No active workbook is open.");
                }

                var availableSheets = new List<string>();
                foreach (Excel.Worksheet s in wb.Worksheets)
                {
                    availableSheets.Add(s.Name);
                }

                bool hasSelection = _app.Selection is Excel.Range;
                var validation = PdfExportEngine.ValidateSettings(settings, availableSheets, hasSelection);
                if (!validation.Success)
                {
                    ShowMessage(validation.Message, "PDF Export Validation", isError: true);
                    return PdfExportResult.Failed(validation.Message);
                }

                using (new ExcelStateScope(_app))
                {
                    var exportedFiles = new List<string>();
                    var quality = settings.Quality == PdfQuality.MinimumSize
                        ? Excel.XlFixedFormatQuality.xlQualityMinimum
                        : Excel.XlFixedFormatQuality.xlQualityStandard;

                    object fromPage = settings.ExportPageRange ? (object)settings.PageFrom : Type.Missing;
                    object toPage = settings.ExportPageRange ? (object)settings.PageTo : Type.Missing;

                    bool needsLayoutOverride = settings.Orientation != PdfPageOrientation.Default ||
                                               settings.PaperSize != PdfPaperSize.Default ||
                                               settings.PageFit != PdfPageFit.Default;

                    if (settings.Scope == PdfExportScope.Selection)
                    {
                        if (!(_app.Selection is Excel.Range selRange))
                        {
                            return PdfExportResult.Failed("The active selection is not a valid cell range.");
                        }

                        string? dir = System.IO.Path.GetDirectoryName(settings.DestinationPath);
                        if (!string.IsNullOrWhiteSpace(dir) && !System.IO.Directory.Exists(dir))
                        {
                            System.IO.Directory.CreateDirectory(dir);
                        }

                        Excel.Worksheet? activeWs = wb.ActiveSheet as Excel.Worksheet;
                        PageSetupState? psState = null;
                        try
                        {
                            if (activeWs != null && needsLayoutOverride)
                            {
                                psState = CapturePageSetup(activeWs.PageSetup);
                                ApplyPageSetupOverrides(activeWs.PageSetup, settings);
                            }

                            selRange.ExportAsFixedFormat(
                                Excel.XlFixedFormatType.xlTypePDF,
                                settings.DestinationPath,
                                quality,
                                settings.IncludeDocumentProperties,
                                settings.IgnorePrintAreas,
                                fromPage,
                                toPage,
                                settings.OpenAfterExport);

                            exportedFiles.Add(settings.DestinationPath);
                        }
                        finally
                        {
                            if (activeWs != null && psState != null)
                            {
                                RestorePageSetup(activeWs.PageSetup, psState);
                            }
                        }
                    }
                    else if (settings.Scope == PdfExportScope.ActiveSheet)
                    {
                        if (!(_app.ActiveSheet is Excel.Worksheet activeWs))
                        {
                            return PdfExportResult.Failed("No active worksheet found.");
                        }

                        string? dir = System.IO.Path.GetDirectoryName(settings.DestinationPath);
                        if (!string.IsNullOrWhiteSpace(dir) && !System.IO.Directory.Exists(dir))
                        {
                            System.IO.Directory.CreateDirectory(dir);
                        }

                        PageSetupState? psState = null;
                        try
                        {
                            if (needsLayoutOverride)
                            {
                                psState = CapturePageSetup(activeWs.PageSetup);
                                ApplyPageSetupOverrides(activeWs.PageSetup, settings);
                            }

                            activeWs.ExportAsFixedFormat(
                                Excel.XlFixedFormatType.xlTypePDF,
                                settings.DestinationPath,
                                quality,
                                settings.IncludeDocumentProperties,
                                settings.IgnorePrintAreas,
                                fromPage,
                                toPage,
                                settings.OpenAfterExport);

                            exportedFiles.Add(settings.DestinationPath);
                        }
                        finally
                        {
                            if (psState != null)
                            {
                                RestorePageSetup(activeWs.PageSetup, psState);
                            }
                        }
                    }
                    else
                    {
                        List<string> targetSheetNames;
                        if (settings.Scope == PdfExportScope.SelectedSheets)
                        {
                            targetSheetNames = settings.SelectedSheets != null ? new List<string>(settings.SelectedSheets) : new List<string>();
                        }
                        else
                        {
                            targetSheetNames = new List<string>();
                            foreach (Excel.Worksheet ws in wb.Worksheets)
                            {
                                if (ws.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                                {
                                    targetSheetNames.Add(ws.Name);
                                }
                            }
                        }

                        if (targetSheetNames.Count == 0)
                        {
                            return PdfExportResult.Failed("No visible worksheets to export.");
                        }

                        if (settings.OutputMode == PdfOutputMode.SeparateFilesPerSheet)
                        {
                            if (!System.IO.Directory.Exists(settings.OutputDirectory))
                            {
                                System.IO.Directory.CreateDirectory(settings.OutputDirectory);
                            }

                            int index = 1;
                            foreach (string sName in targetSheetNames)
                            {
                                Excel.Worksheet? ws = null;
                                foreach (Excel.Worksheet w in wb.Worksheets)
                                {
                                    if (string.Equals(w.Name, sName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        ws = w;
                                        break;
                                    }
                                }

                                if (ws == null) continue;

                                string fileName = PdfExportEngine.ResolveFileName(
                                    settings.FileNamePattern,
                                    wb.Name,
                                    ws.Name,
                                    index++,
                                    DateTime.Now);
                                string fullPath = System.IO.Path.Combine(settings.OutputDirectory, fileName);

                                PageSetupState? psState = null;
                                try
                                {
                                    if (needsLayoutOverride)
                                    {
                                        psState = CapturePageSetup(ws.PageSetup);
                                        ApplyPageSetupOverrides(ws.PageSetup, settings);
                                    }

                                    ws.ExportAsFixedFormat(
                                        Excel.XlFixedFormatType.xlTypePDF,
                                        fullPath,
                                        quality,
                                        settings.IncludeDocumentProperties,
                                        settings.IgnorePrintAreas,
                                        fromPage,
                                        toPage,
                                        OpenAfterPublish: false);

                                    exportedFiles.Add(fullPath);
                                }
                                finally
                                {
                                    if (psState != null)
                                    {
                                        RestorePageSetup(ws.PageSetup, psState);
                                    }
                                }
                            }

                            if (exportedFiles.Count == 0)
                            {
                                return PdfExportResult.Failed("No worksheets were successfully exported.");
                            }

                            if (settings.OpenAfterExport && exportedFiles.Count > 0)
                            {
                                try
                                {
                                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                    {
                                        FileName = exportedFiles[0],
                                        UseShellExecute = true
                                    });
                                }
                                catch { }
                            }
                        }
                        else
                        {
                            // Single Combined PDF
                            string? dir = System.IO.Path.GetDirectoryName(settings.DestinationPath);
                            if (!string.IsNullOrWhiteSpace(dir) && !System.IO.Directory.Exists(dir))
                            {
                                System.IO.Directory.CreateDirectory(dir);
                            }

                            if (targetSheetNames.Count == 1)
                            {
                                Excel.Worksheet? ws = null;
                                foreach (Excel.Worksheet w in wb.Worksheets)
                                {
                                    if (string.Equals(w.Name, targetSheetNames[0], StringComparison.OrdinalIgnoreCase))
                                    {
                                        ws = w;
                                        break;
                                    }
                                }

                                if (ws == null) return PdfExportResult.Failed($"Worksheet '{targetSheetNames[0]}' not found.");

                                PageSetupState? psState = null;
                                try
                                {
                                    if (needsLayoutOverride)
                                    {
                                        psState = CapturePageSetup(ws.PageSetup);
                                        ApplyPageSetupOverrides(ws.PageSetup, settings);
                                    }

                                    ws.ExportAsFixedFormat(
                                        Excel.XlFixedFormatType.xlTypePDF,
                                        settings.DestinationPath,
                                        quality,
                                        settings.IncludeDocumentProperties,
                                        settings.IgnorePrintAreas,
                                        fromPage,
                                        toPage,
                                        settings.OpenAfterExport);

                                    exportedFiles.Add(settings.DestinationPath);
                                }
                                finally
                                {
                                    if (psState != null)
                                    {
                                        RestorePageSetup(ws.PageSetup, psState);
                                    }
                                }
                            }
                            else if (settings.Scope == PdfExportScope.EntireWorkbook && !needsLayoutOverride)
                            {
                                wb.ExportAsFixedFormat(
                                    Excel.XlFixedFormatType.xlTypePDF,
                                    settings.DestinationPath,
                                    quality,
                                    settings.IncludeDocumentProperties,
                                    settings.IgnorePrintAreas,
                                    fromPage,
                                    toPage,
                                    settings.OpenAfterExport);

                                exportedFiles.Add(settings.DestinationPath);
                            }
                            else
                            {
                                var originalActive = wb.ActiveSheet as Excel.Worksheet;
                                var restoredStates = new List<(Excel.Worksheet ws, PageSetupState state)>();

                                try
                                {
                                    if (needsLayoutOverride)
                                    {
                                        foreach (var sName in targetSheetNames)
                                        {
                                            foreach (Excel.Worksheet w in wb.Worksheets)
                                            {
                                                if (string.Equals(w.Name, sName, StringComparison.OrdinalIgnoreCase))
                                                {
                                                    restoredStates.Add((w, CapturePageSetup(w.PageSetup)));
                                                    ApplyPageSetupOverrides(w.PageSetup, settings);
                                                    break;
                                                }
                                            }
                                        }
                                    }

                                    bool isFirst = true;
                                    foreach (var sName in targetSheetNames)
                                    {
                                        foreach (Excel.Worksheet w in wb.Worksheets)
                                        {
                                            if (string.Equals(w.Name, sName, StringComparison.OrdinalIgnoreCase) &&
                                                w.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                                            {
                                                w.Select(Replace: isFirst);
                                                isFirst = false;
                                                break;
                                            }
                                        }
                                    }

                                    dynamic? selSheets = _app.ActiveWindow?.SelectedSheets;
                                    if (selSheets != null)
                                    {
                                        selSheets.ExportAsFixedFormat(
                                            Excel.XlFixedFormatType.xlTypePDF,
                                            settings.DestinationPath,
                                            quality,
                                            settings.IncludeDocumentProperties,
                                            settings.IgnorePrintAreas,
                                            fromPage,
                                            toPage,
                                            settings.OpenAfterExport);
                                    }
                                    else if (wb.ActiveSheet is Excel.Worksheet grpWs)
                                    {
                                        grpWs.ExportAsFixedFormat(
                                            Excel.XlFixedFormatType.xlTypePDF,
                                            settings.DestinationPath,
                                            quality,
                                            settings.IncludeDocumentProperties,
                                            settings.IgnorePrintAreas,
                                            fromPage,
                                            toPage,
                                            settings.OpenAfterExport);
                                    }
                                    else
                                    {
                                        wb.ExportAsFixedFormat(
                                            Excel.XlFixedFormatType.xlTypePDF,
                                            settings.DestinationPath,
                                            quality,
                                            settings.IncludeDocumentProperties,
                                            settings.IgnorePrintAreas,
                                            fromPage,
                                            toPage,
                                            settings.OpenAfterExport);
                                    }

                                    exportedFiles.Add(settings.DestinationPath);
                                }
                                finally
                                {
                                    foreach (var item in restoredStates)
                                    {
                                        RestorePageSetup(item.ws.PageSetup, item.state);
                                    }
                                    try { originalActive?.Select(); } catch { }
                                }
                            }
                        }
                    }

                    if (settings.OpenOutputFolder)
                    {
                        string folder = settings.OutputMode == PdfOutputMode.SeparateFilesPerSheet
                            ? settings.OutputDirectory
                            : System.IO.Path.GetDirectoryName(settings.DestinationPath) ?? "";

                        if (!string.IsNullOrWhiteSpace(folder) && System.IO.Directory.Exists(folder))
                        {
                            try
                            {
                                if (settings.OutputMode == PdfOutputMode.SingleCombinedFile && System.IO.File.Exists(settings.DestinationPath))
                                {
                                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{settings.DestinationPath}\"");
                                }
                                else
                                {
                                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                    {
                                        FileName = folder,
                                        UseShellExecute = true
                                    });
                                }
                            }
                            catch { }
                        }
                    }

                    string successMsg = exportedFiles.Count == 1
                        ? $"PDF published successfully: '{exportedFiles[0]}'."
                        : $"Exported {exportedFiles.Count} PDF files to '{settings.OutputDirectory}'.";

                    return PdfExportResult.Succeeded(exportedFiles.Count, exportedFiles, successMsg);
                }
            }
            catch (COMException comEx)
            {
                string msg = comEx.Message;
                if (comEx.ErrorCode == unchecked((int)0x800A03EC) || msg.Contains("sharing violation") || msg.Contains("used by another process") || msg.Contains("access"))
                {
                    msg = "The destination PDF file may be open or locked in another program (such as Adobe Acrobat or a web browser). Please close the file and try again.";
                }
                ShowMessage(msg, "PDF Export Error", isError: true);
                return PdfExportResult.Failed(msg);
            }
            catch (Exception ex)
            {
                ShowMessage($"Error exporting PDF: {ex.Message}", "PDF Export Error", isError: true);
                return PdfExportResult.Failed(ex.Message);
            }
        }

        public int BatchConvertWorkbooks(IEnumerable<string> filePaths, string outputDirectory, string targetExtension)
        {
            if (filePaths == null) throw new ArgumentNullException(nameof(filePaths));
            if (string.IsNullOrWhiteSpace(outputDirectory)) throw new ArgumentNullException(nameof(outputDirectory));
            if (string.IsNullOrWhiteSpace(targetExtension)) throw new ArgumentNullException(nameof(targetExtension));

            if (!System.IO.Directory.Exists(outputDirectory))
            {
                System.IO.Directory.CreateDirectory(outputDirectory);
            }

            string cleanExt = targetExtension.TrimStart('.').ToLowerInvariant();
            int count = 0;

            using (new ExcelStateScope(_app))
            {
                foreach (string filePath in filePaths)
                {
                    if (!System.IO.File.Exists(filePath)) continue;

                    Excel.Workbook? srcWb = null;
                    try
                    {
                        srcWb = _app.Workbooks.Open(filePath, ReadOnly: true);
                        string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath);
                        string outPath = System.IO.Path.Combine(outputDirectory, $"{fileNameWithoutExt}.{cleanExt}");

                        if (string.Equals(cleanExt, "pdf", StringComparison.OrdinalIgnoreCase))
                        {
                            srcWb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outPath);
                        }
                        else if (string.Equals(cleanExt, "csv", StringComparison.OrdinalIgnoreCase))
                        {
                            srcWb.SaveAs(outPath, Excel.XlFileFormat.xlCSV);
                        }
                        else if (string.Equals(cleanExt, "xls", StringComparison.OrdinalIgnoreCase))
                        {
                            srcWb.SaveAs(outPath, Excel.XlFileFormat.xlExcel8);
                        }
                        else
                        {
                            srcWb.SaveAs(outPath, Excel.XlFileFormat.xlOpenXMLWorkbook);
                        }

                        count++;
                    }
                    catch
                    {
                        // Continue processing remaining files
                    }
                    finally
                    {
                        if (srcWb != null)
                        {
                            try { srcWb.Close(SaveChanges: false); } catch { }
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(srcWb); } catch { }
                        }
                    }
                }
            }

            return count;
        }

        public IReadOnlyList<SheetProtectionStatus> GetSheetProtectionStatuses()
        {
            var list = new List<SheetProtectionStatus>();
            if (!HasActiveWorkbook) return list;
            try
            {
                foreach (Excel.Worksheet ws in _app.ActiveWorkbook.Worksheets)
                {
                    bool isProt = false;
                    try { isProt = ws.ProtectContents; } catch { }
                    list.Add(new SheetProtectionStatus(ws.Name, isProt));
                }
            }
            catch { }
            return list;
        }

        public (int successCount, List<string> errors) BatchProtectWorksheets(IEnumerable<string> sheetNames, string? password, SheetProtectionOptions options)
        {
            var errors = new List<string>();
            int successCount = 0;
            if (!HasActiveWorkbook) return (0, errors);

            using (new ExcelStateScope(_app))
            {
                foreach (string name in sheetNames)
                {
                    try
                    {
                        var ws = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[name];
                        object pwd = string.IsNullOrEmpty(password) ? Type.Missing : (object)password!;
                        ws.Protect(
                            Password: pwd,
                            DrawingObjects: true,
                            Contents: true,
                            Scenarios: true,
                            UserInterfaceOnly: Type.Missing,
                            AllowFormattingCells: options.FormatCells,
                            AllowFormattingColumns: options.FormatColumns,
                            AllowFormattingRows: options.FormatRows,
                            AllowInsertingColumns: options.InsertColumns,
                            AllowInsertingRows: options.InsertRows,
                            AllowInsertingHyperlinks: options.InsertHyperlinks,
                            AllowDeletingColumns: options.DeleteColumns,
                            AllowDeletingRows: options.DeleteRows,
                            AllowSorting: options.Sort,
                            AllowFiltering: options.AutoFilter,
                            AllowUsingPivotTables: options.UsePivotTables);

                        if (!options.SelectLockedCells)
                            ws.EnableSelection = Excel.XlEnableSelection.xlNoSelection;
                        else if (!options.SelectUnlockedCells)
                            ws.EnableSelection = Excel.XlEnableSelection.xlUnlockedCells;
                        else
                            ws.EnableSelection = Excel.XlEnableSelection.xlNoRestrictions;

                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{name}: {ex.Message}");
                    }
                }
            }
            return (successCount, errors);
        }

        public (int successCount, List<string> errors) BatchUnprotectWorksheets(IEnumerable<string> sheetNames, string? password)
        {
            var errors = new List<string>();
            int successCount = 0;
            if (!HasActiveWorkbook) return (0, errors);

            using (new ExcelStateScope(_app))
            {
                foreach (string name in sheetNames)
                {
                    try
                    {
                        var ws = (Excel.Worksheet)_app.ActiveWorkbook.Worksheets[name];
                        if (string.IsNullOrEmpty(password))
                        {
                            ws.Unprotect();
                        }
                        else
                        {
                            ws.Unprotect(password);
                        }
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{name}: {ex.Message}");
                    }
                }
            }
            return (successCount, errors);
        }

        public bool SaveActiveWorkbookCopy(string destinationPath)
        {
            if (!HasActiveWorkbook) return false;
            try
            {
                string? dir = System.IO.Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(dir) && !System.IO.Directory.Exists(dir))
                {
                    System.IO.Directory.CreateDirectory(dir);
                }
                _app.ActiveWorkbook.SaveCopyAs(destinationPath);
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to save workbook copy: {ex.Message}", "Snapshot Error", isError: true);
                return false;
            }
        }

        public bool RestoreWorkbookFromSnapshot(string snapshotPath)
        {
            if (!System.IO.File.Exists(snapshotPath)) return false;
            try
            {
                if (HasActiveWorkbook)
                {
                    _app.ActiveWorkbook.Close(SaveChanges: false);
                }
                _app.Workbooks.Open(snapshotPath);
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to restore snapshot: {ex.Message}", "Restore Error", isError: true);
                return false;
            }
        }

        public bool CreateTargetVsActualChart(string title, int startRow, int startCol, int rowCount, int colCount, int categoryColIndex, int actualColIndex, int targetColIndex)
        {
            if (!HasActiveWorksheet) return false;
            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                var chartObjects = (Excel.ChartObjects)ws.ChartObjects();

                var cell1 = (Excel.Range)ws.Cells[startRow, startCol];
                var cell2 = (Excel.Range)ws.Cells[startRow + rowCount - 1, startCol + colCount - 1];
                var range = ws.Range[cell1, cell2];

                double left = 100;
                double top = 100;
                try
                {
                    left = (double)range.Left + (double)range.Width + 20;
                    top = (double)range.Top;
                }
                catch { }

                var chartObject = chartObjects.Add(left, top, 520, 320);
                var chart = chartObject.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.SetSourceData(range, Excel.XlRowCol.xlColumns);

                try
                {
                    var seriesCol = (Excel.SeriesCollection)chart.SeriesCollection();
                    if (seriesCol.Count >= 2)
                    {
                        var targetSeries = (Excel.Series)seriesCol.Item(2);
                        targetSeries.ChartType = Excel.XlChartType.xlLineMarkers;
                    }
                }
                catch { }

                chart.HasTitle = true;
                chart.ChartTitle.Text = string.IsNullOrWhiteSpace(title) ? "Target vs. Actual Performance" : title;

                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to create chart: {ex.Message}", "Chart Error", isError: true);
                return false;
            }
        }

        public IReadOnlyList<NamedRangeInfo> GetNamedRanges()
        {
            var list = new List<NamedRangeInfo>();
            if (!HasActiveWorkbook) return list;
            try
            {
                foreach (Excel.Name name in _app.ActiveWorkbook.Names)
                {
                    try
                    {
                        string n = name.Name;
                        string refersTo = name.RefersTo?.ToString() ?? "";
                        string comment = "";
                        try { comment = name.Comment ?? ""; } catch { }

                        string scope = "Workbook";
                        try
                        {
                            if (name.Parent is Excel.Worksheet ws)
                            {
                                scope = ws.Name;
                            }
                        }
                        catch { }

                        bool isValid = !refersTo.Contains("#REF!");
                        string targetAddress = refersTo.TrimStart('=');
                        list.Add(new NamedRangeInfo(n, refersTo, scope, comment, isValid, targetAddress));
                    }
                    catch { }
                }
            }
            catch { }
            return list;
        }

        public bool GoToNamedRange(string name)
        {
            if (!HasActiveWorkbook || string.IsNullOrWhiteSpace(name)) return false;
            try
            {
                _app.Goto(name, true);
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not navigate to named range '{name}': {ex.Message}", "Navigation Error", isError: true);
                return false;
            }
        }

        public bool DeleteNamedRange(string name)
        {
            if (!HasActiveWorkbook || string.IsNullOrWhiteSpace(name)) return false;
            try
            {
                var n = _app.ActiveWorkbook.Names.Item(name);
                n.Delete();
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Could not delete named range '{name}': {ex.Message}", "Delete Error", isError: true);
                return false;
            }
        }

        public IReadOnlyList<WorksheetColumnInfo> GetWorksheetColumns(int maxCols = 100)
        {
            var list = new List<WorksheetColumnInfo>();
            if (!HasActiveWorksheet) return list;
            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                var used = ws.UsedRange;
                int startCol = 1;
                int endCol = 1;
                try
                {
                    startCol = used.Column;
                    endCol = startCol + used.Columns.Count - 1;
                }
                catch { }

                if (endCol > maxCols) endCol = maxCols;

                for (int col = startCol; col <= endCol; col++)
                {
                    string colLetter = ColumnLetterConverter.IndexToLetter(col);
                    string header = "";
                    try
                    {
                        var cell = (Excel.Range)ws.Cells[1, col];
                        header = cell.Value2?.ToString()?.Trim() ?? "";
                    }
                    catch { }

                    double width = 8.43;
                    bool isHidden = false;
                    try
                    {
                        var colRange = (Excel.Range)ws.Columns[col];
                        width = Convert.ToDouble(colRange.ColumnWidth);
                        isHidden = (bool)colRange.Hidden;
                    }
                    catch { }

                    list.Add(new WorksheetColumnInfo(col, colLetter, header, width, isHidden));
                }
            }
            catch { }
            return list;
        }

        public void SelectColumn(int columnIndex)
        {
            if (!HasActiveWorksheet || columnIndex < 1) return;
            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                ((Excel.Range)ws.Columns[columnIndex]).Select();
            }
            catch { }
        }

        public void SetColumnsVisibility(IEnumerable<int> columnIndices, bool isHidden)
        {
            if (!HasActiveWorksheet) return;
            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                foreach (int c in columnIndices)
                {
                    try
                    {
                        ((Excel.Range)ws.Columns[c]).Hidden = isHidden;
                    }
                    catch { }
                }
            }
            catch { }
        }

        public void AutoFitColumns(IEnumerable<int> columnIndices)
        {
            if (!HasActiveWorksheet) return;
            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                foreach (int c in columnIndices)
                {
                    try
                    {
                        ((Excel.Range)ws.Columns[c]).AutoFit();
                    }
                    catch { }
                }
            }
            catch { }
        }

        public CellContentDetails? GetActiveCellDetails()
        {
            if (!HasActiveWorksheet) return null;
            try
            {
                var cell = _app.ActiveCell;
                if (cell == null) return null;

                string address = cell.get_Address();
                string sheetName = ((Excel.Worksheet)_app.ActiveSheet).Name;
                string valText = cell.Value2?.ToString() ?? "";
                string formulaText = "";
                try { formulaText = cell.Formula?.ToString() ?? ""; } catch { }
                string numFmt = "";
                try { numFmt = cell.NumberFormat?.ToString() ?? ""; } catch { }

                string dataType = "String";
                object? raw = cell.Value2;
                if (raw == null) dataType = "Empty / Blank";
                else if (raw is double || raw is int || raw is float || raw is decimal) dataType = "Numeric";
                else if (raw is bool) dataType = "Boolean";
                else if (valText.StartsWith("#")) dataType = "Error";
                else if (raw is DateTime) dataType = "Date";

                return new CellContentDetails(address, sheetName, valText, formulaText, numFmt, dataType);
            }
            catch
            {
                return null;
            }
        }

        public (int row, int col, string rowHeader, string colHeader, string address) GetActiveCellContext()
        {
            if (!HasActiveWorksheet) return (1, 1, "", "", "$A$1");
            try
            {
                var cell = _app.ActiveCell;
                if (cell == null) return (1, 1, "", "", "$A$1");

                int row = cell.Row;
                int col = cell.Column;
                string addr = cell.get_Address();
                var ws = (Excel.Worksheet)_app.ActiveSheet;

                string colHeader = "";
                try
                {
                    colHeader = ((Excel.Range)ws.Cells[1, col]).Value2?.ToString()?.Trim() ?? "";
                }
                catch { }

                string rowHeader = "";
                try
                {
                    rowHeader = ((Excel.Range)ws.Cells[row, 1]).Value2?.ToString()?.Trim() ?? "";
                }
                catch { }

                return (row, col, rowHeader, colHeader, addr);
            }
            catch
            {
                return (1, 1, "", "", "$A$1");
            }
        }

        public void HighlightActiveRowAndColumn(int row, int col)
        {
            if (!HasActiveWorksheet) return;
            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                var rowRange = (Excel.Range)ws.Rows[row];
                var colRange = (Excel.Range)ws.Columns[col];
                var unionRange = _app.Union(rowRange, colRange);
                unionRange.Select();
            }
            catch { }
        }

        public TableDataModel? GetActiveTableData(int maxSampleRows = 500)
        {
            if (!HasActiveWorksheet) return null;
            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                Excel.Range targetRange = _app.Selection as Excel.Range ?? ws.UsedRange;
                if (targetRange == null) return null;

                int startRow = targetRange.Row;
                int startCol = targetRange.Column;
                int rowCount = targetRange.Rows.Count;
                int colCount = targetRange.Columns.Count;

                if (rowCount < 1 || colCount < 1) return null;

                var headers = new List<string>();
                for (int c = 1; c <= colCount; c++)
                {
                    var cell = (Excel.Range)ws.Cells[startRow, startCol + c - 1];
                    string h = cell.Value2?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrWhiteSpace(h)) h = ColumnLetterConverter.IndexToLetter(startCol + c - 1);
                    headers.Add(h);
                }

                var rows = new List<object?[]>();
                int dataRows = Math.Min(rowCount - 1, maxSampleRows);
                if (dataRows > 0)
                {
                    var dataRange = ws.Range[ws.Cells[startRow + 1, startCol], ws.Cells[startRow + dataRows, startCol + colCount - 1]];
                    object[,] vals = (object[,])dataRange.Value2;
                    if (vals != null)
                    {
                        int rStart = vals.GetLowerBound(0);
                        int rEnd = vals.GetUpperBound(0);
                        int cStart = vals.GetLowerBound(1);

                        for (int r = rStart; r <= rEnd; r++)
                        {
                            var rowVals = new object?[colCount];
                            for (int c = 0; c < colCount; c++)
                            {
                                rowVals[c] = vals[r, cStart + c];
                            }
                            rows.Add(rowVals);
                        }
                    }
                }

                return new TableDataModel(headers, rows, startRow, startCol);
            }
            catch
            {
                return null;
            }
        }

        public bool AppendTableRow(object?[] rowData, int tableStartRow, int tableStartCol, int colCount)
        {
            if (!HasActiveWorksheet || rowData == null || rowData.Length == 0) return false;
            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                int lastRow = ws.Cells[ws.Rows.Count, tableStartCol].End(Excel.XlDirection.xlUp).Row;
                int newRow = Math.Max(tableStartRow + 1, lastRow + 1);

                var destRange = ws.Range[ws.Cells[newRow, tableStartCol], ws.Cells[newRow, tableStartCol + colCount - 1]];
                var matrix = new object[1, colCount];
                for (int c = 0; c < colCount; c++)
                {
                    matrix[0, c] = (c < rowData.Length && rowData[c] != null) ? rowData[c]! : "";
                }
                destRange.Value2 = matrix;
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to append record: {ex.Message}", "Append Error", isError: true);
                return false;
            }
        }

        public void ShowMessage(string message, string title, bool isError = false)
        {
            MessageBox.Show(message, title, MessageBoxButton.OK, isError ? MessageBoxImage.Error : MessageBoxImage.Information);
        }

        public IReadOnlyList<CellReferenceInfo> GetActiveCellPrecedents(bool directOnly = true)
        {
            var list = new List<CellReferenceInfo>();
            if (!HasActiveWorksheet) return list;

            try
            {
                var cell = _app.ActiveCell;
                if (cell == null) return list;

                string activeSheetName = ((Excel.Worksheet)_app.ActiveSheet).Name;
                Excel.Range? precRange = null;

                try
                {
                    precRange = directOnly ? cell.DirectPrecedents : cell.Precedents;
                }
                catch (COMException)
                {
                    // 0x800A03EC thrown by Excel when no precedent cells exist
                    precRange = null;
                }

                if (precRange != null)
                {
                    foreach (Excel.Range area in precRange.Areas)
                    {
                        int count = area.CountLarge > 200 ? 200 : area.Count;
                        foreach (Excel.Range pCell in area)
                        {
                            string addr = pCell.get_Address();
                            string sName = ((Excel.Worksheet)pCell.Worksheet).Name;
                            string val = pCell.Value2?.ToString() ?? "";
                            string formula = "";
                            try { formula = pCell.Formula?.ToString() ?? ""; } catch { }

                            list.Add(new CellReferenceInfo(sName, addr, val, formula, directOnly));
                            if (list.Count >= 300) break;
                        }
                        if (list.Count >= 300) break;
                    }
                }

                // If formula references other sheets (not supported by .Precedents COM property), scan formula string
                string activeFormula = "";
                try { activeFormula = cell.Formula?.ToString() ?? ""; } catch { }

                if (!string.IsNullOrWhiteSpace(activeFormula) && activeFormula.Contains("!"))
                {
                    var regex = new Regex(@"(?<sheet>'(?:[^']|'')+'|[A-Za-z0-9_]+)!(?<ref>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?)");
                    var matches = regex.Matches(activeFormula);
                    var wb = _app.ActiveWorkbook;

                    foreach (Match m in matches)
                    {
                        string targetSheet = m.Groups["sheet"].Value.Trim('\'');
                        string targetRef = m.Groups["ref"].Value;

                        if (!targetSheet.Equals(activeSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            try
                            {
                                var ws = (Excel.Worksheet)wb.Worksheets[targetSheet];
                                var extRange = ws.Range[targetRef];
                                string val = extRange.Value2?.ToString() ?? "";
                                string fText = "";
                                try { fText = extRange.Formula?.ToString() ?? ""; } catch { }

                                list.Add(new CellReferenceInfo(targetSheet, targetRef, val, fText, directOnly));
                            }
                            catch { }
                        }
                    }
                }
            }
            catch { }

            return list;
        }

        public IReadOnlyList<CellReferenceInfo> GetActiveCellDependents(bool directOnly = true)
        {
            var list = new List<CellReferenceInfo>();
            if (!HasActiveWorksheet) return list;

            try
            {
                var cell = _app.ActiveCell;
                if (cell == null) return list;

                Excel.Range? depRange = null;
                try
                {
                    depRange = directOnly ? cell.DirectDependents : cell.Dependents;
                }
                catch (COMException)
                {
                    // 0x800A03EC thrown when no dependent cells exist
                    depRange = null;
                }

                if (depRange != null)
                {
                    foreach (Excel.Range area in depRange.Areas)
                    {
                        int count = area.CountLarge > 200 ? 200 : area.Count;
                        foreach (Excel.Range dCell in area)
                        {
                            string addr = dCell.get_Address();
                            string sName = ((Excel.Worksheet)dCell.Worksheet).Name;
                            string val = dCell.Value2?.ToString() ?? "";
                            string formula = "";
                            try { formula = dCell.Formula?.ToString() ?? ""; } catch { }

                            list.Add(new CellReferenceInfo(sName, addr, val, formula, directOnly));
                            if (list.Count >= 300) break;
                        }
                        if (list.Count >= 300) break;
                    }
                }
            }
            catch { }

            return list;
        }

        public bool SelectCellAddress(string sheetName, string cellAddress)
        {
            if (!HasActiveWorkbook || string.IsNullOrWhiteSpace(cellAddress)) return false;

            try
            {
                var wb = _app.ActiveWorkbook;
                Excel.Worksheet ws;
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    try
                    {
                        ws = (Excel.Worksheet)wb.Worksheets[sheetName];
                    }
                    catch
                    {
                        ws = (Excel.Worksheet)_app.ActiveSheet;
                    }
                }
                else
                {
                    ws = (Excel.Worksheet)_app.ActiveSheet;
                }

                ws.Activate();
                Excel.Range target = ws.Range[cellAddress];
                target.Select();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void ApplyNumberFormat(string formatString, bool selectionOnly = true)
        {
            if (!HasActiveWorksheet || string.IsNullOrEmpty(formatString)) return;

            try
            {
                using (new ExcelStateScope(_app))
                {
                    var ws = (Excel.Worksheet)_app.ActiveSheet;
                    Excel.Range target = selectionOnly ? (_app.Selection as Excel.Range ?? ws.UsedRange) : ws.UsedRange;
                    if (target != null)
                    {
                        target.NumberFormat = formatString;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to apply number format: {ex.Message}", "Formatting Error", isError: true);
            }
        }

        public string? GetActiveCellNumberFormat()
        {
            if (!HasActiveWorksheet) return null;
            try
            {
                var cell = _app.ActiveCell;
                return cell?.NumberFormat?.ToString();
            }
            catch
            {
                return null;
            }
        }

        public void InsertFormulaAtActiveCell(string formula)
        {
            if (!HasActiveWorksheet || string.IsNullOrWhiteSpace(formula)) return;

            try
            {
                var cell = _app.ActiveCell;
                if (cell != null)
                {
                    cell.Formula = formula;
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to insert formula: {ex.Message}", "Formula Insertion Error", isError: true);
            }
        }

        public IReadOnlyList<object?[,]> GetMultiAreaSelectedValues()
        {
            var list = new List<object?[,]>();
            if (!HasActiveWorksheet) return list;

            try
            {
                if (_app.Selection is Excel.Range sel)
                {
                    foreach (Excel.Range area in sel.Areas)
                    {
                        object?[,] vals = area.Rows.Count == 1 && area.Columns.Count == 1
                            ? new object?[,] { { area.Value2 } }
                            : (object?[,])area.Value2;

                        if (vals != null) list.Add(vals);
                    }
                }
            }
            catch { }
            return list;
        }

        public IReadOnlyList<string> GetMultiAreaAddresses()
        {
            var list = new List<string>();
            if (!HasActiveWorksheet) return list;

            try
            {
                if (_app.Selection is Excel.Range sel)
                {
                    foreach (Excel.Range area in sel.Areas)
                    {
                        list.Add(area.get_Address());
                    }
                }
            }
            catch { }
            return list;
        }

        public void SetMultiAreaSelectedValues(IReadOnlyList<object?[,]> areaValues)
        {
            if (areaValues == null) throw new ArgumentNullException(nameof(areaValues));
            if (!(_app.Selection is Excel.Range range)) return;

            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before modifying cell values.", "Sheet Protected", isError: true);
                return;
            }

            using (new ExcelStateScope(_app))
            {
                int count = Math.Min(range.Areas.Count, areaValues.Count);
                for (int i = 1; i <= count; i++)
                {
                    Excel.Range area = range.Areas[i];
                    var matrix = areaValues[i - 1];
                    if (matrix != null)
                    {
                        RecordBeforeRangeChange(area, matrix, isFormula: false, customDesc: "Multi-Area Cell Value Update");
                        if (area.Rows.Count == 1 && area.Columns.Count == 1)
                        {
                            int r0 = matrix.GetLowerBound(0);
                            int c0 = matrix.GetLowerBound(1);
                            area.Value2 = matrix[r0, c0];
                        }
                        else
                        {
                            area.Value2 = matrix;
                        }
                    }
                }
            }
        }

        public bool WriteMatrixToAddress(string targetAddress, object?[,] matrix)
        {
            if (!HasActiveWorksheet || string.IsNullOrWhiteSpace(targetAddress) || matrix == null) return false;

            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                var originCell = ws.Range[targetAddress];
                int startR = originCell.Row;
                int startC = originCell.Column;

                int rCount = matrix.GetUpperBound(0) - matrix.GetLowerBound(0) + 1;
                int cCount = matrix.GetUpperBound(1) - matrix.GetLowerBound(1) + 1;

                using (new ExcelStateScope(_app))
                {
                    var destRange = ws.Range[ws.Cells[startR, startC], ws.Cells[startR + rCount - 1, startC + cCount - 1]];
                    destRange.Value2 = matrix;
                }
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to write data to {targetAddress}: {ex.Message}", "Paste Error", isError: true);
                return false;
            }
        }

        public IReadOnlyList<int> GetVisibleRowIndices(int startRow, int maxRows = 2000)
        {
            var list = new List<int>();
            if (!HasActiveWorksheet || startRow < 1) return list;

            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                int maxAvailable = ws.Rows.Count;
                int r = startRow;

                while (list.Count < maxRows && r <= maxAvailable)
                {
                    bool isHidden = false;
                    try
                    {
                        isHidden = (bool)((Excel.Range)ws.Rows[r]).Hidden;
                    }
                    catch { }

                    if (!isHidden)
                    {
                        list.Add(r);
                    }
                    r++;
                }
            }
            catch { }
            return list;
        }

        public bool PasteIntoVisibleCells(object?[,] sourceMatrix, int startRow, int startCol)
        {
            if (!HasActiveWorksheet || sourceMatrix == null) return false;

            try
            {
                int rCount = sourceMatrix.GetUpperBound(0) - sourceMatrix.GetLowerBound(0) + 1;
                int cCount = sourceMatrix.GetUpperBound(1) - sourceMatrix.GetLowerBound(1) + 1;

                var visibleRows = GetVisibleRowIndices(startRow, rCount);
                if (visibleRows.Count == 0) return false;

                var mappings = VisibleCellsPasteEngine.ComputeMapping(sourceMatrix, visibleRows);
                var ws = (Excel.Worksheet)_app.ActiveSheet;

                using (new ExcelStateScope(_app))
                {
                    foreach (var m in mappings)
                    {
                        var rowMatrix = new object[1, cCount];
                        for (int c = 0; c < cCount; c++)
                        {
                            rowMatrix[0, c] = m.Values[c]!;
                        }

                        var dest = ws.Range[ws.Cells[m.TargetRowIndex, startCol], ws.Cells[m.TargetRowIndex, startCol + cCount - 1]];
                        dest.Value2 = rowMatrix;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to paste into visible cells: {ex.Message}", "Paste Error", isError: true);
                return false;
            }
        }

        public IReadOnlyList<MergedAreaInfo> GetMergedAreasInSelection()
        {
            var list = new List<MergedAreaInfo>();
            if (!HasActiveWorksheet) return list;

            try
            {
                if (_app.Selection is Excel.Range sel)
                {
                    var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    // Quick safeguard: if whole selection is unmerged, exit fast
                    if (sel.MergeCells is bool b && !b) return list;

                    foreach (Excel.Range cell in sel.Cells)
                    {
                        if (cell.MergeCells is bool isMerged && isMerged)
                        {
                            var mArea = cell.MergeArea;
                            string addr = mArea.get_Address();
                            if (!visited.Contains(addr))
                            {
                                visited.Add(addr);
                                int sRow = mArea.Row;
                                int sCol = mArea.Column;
                                int rCount = mArea.Rows.Count;
                                int cCount = mArea.Columns.Count;
                                object? val = ((Excel.Range)mArea.Cells[1, 1]).Value2;
                                list.Add(new MergedAreaInfo(sRow, sCol, rCount, cCount, addr, val));
                            }
                        }
                    }
                }
            }
            catch { }
            return list;
        }

        public int UnmergeAndFillSelection()
        {
            if (!HasActiveWorksheet) return 0;

            try
            {
                var mergedAreas = GetMergedAreasInSelection();
                if (mergedAreas.Count == 0) return 0;

                var ws = (Excel.Worksheet)_app.ActiveSheet;
                int totalCellsFilled = 0;

                using (new ExcelStateScope(_app))
                {
                    foreach (var area in mergedAreas)
                    {
                        object? val = area.Value;
                        var targetRange = ws.Range[area.Address];
                        targetRange.UnMerge();
                        targetRange.Value2 = val;
                        totalCellsFilled += (area.RowCount * area.ColCount);
                    }
                }
                return totalCellsFilled;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to unmerge and fill: {ex.Message}", "Unmerge Error", isError: true);
                return 0;
            }
        }

        public bool WriteMatrixToNewWorksheet(string sheetName, object?[,] matrix)
        {
            if (!HasActiveWorkbook || matrix == null) return false;

            try
            {
                var wb = _app.ActiveWorkbook;
                var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Excel.Worksheet w in wb.Worksheets)
                {
                    existing.Add(w.Name);
                }

                string validName = string.IsNullOrWhiteSpace(sheetName) ? "SampledData" : sheetName;
                if (existing.Contains(validName) || validName.Length > 31)
                {
                    string baseName = validName.Length > 25 ? validName.Substring(0, 25) : validName;
                    int counter = 1;
                    do
                    {
                        validName = $"{baseName}_{counter++}";
                    } while (existing.Contains(validName));
                }

                var ws = (Excel.Worksheet)wb.Worksheets.Add();
                try { ws.Name = validName; } catch { }

                int rCount = matrix.GetUpperBound(0) - matrix.GetLowerBound(0) + 1;
                int cCount = matrix.GetUpperBound(1) - matrix.GetLowerBound(1) + 1;

                using (new ExcelStateScope(_app))
                {
                    var destRange = ws.Range[ws.Cells[1, 1], ws.Cells[rCount, cCount]];
                    destRange.Value2 = matrix;
                }

                ws.Activate();
                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to output to new worksheet: {ex.Message}", "Worksheet Error", isError: true);
                return false;
            }
        }

        public IReadOnlyList<FormatMatchItem> FindCellsByFormat(CellFormatCriteria criteria, bool selectionOnly = false)
        {
            var matches = new List<FormatMatchItem>();
            if (!HasActiveWorksheet || criteria == null) return matches;

            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                Excel.Range scanRange = selectionOnly ? (_app.Selection as Excel.Range ?? ws.UsedRange) : ws.UsedRange;
                if (scanRange == null) return matches;

                string sheetName = ws.Name;
                int scanned = 0;

                foreach (Excel.Range cell in scanRange.Cells)
                {
                    scanned++;
                    if (scanned > 5000) break; // Performance safeguard

                    bool isMatch = true;

                    // Fill Color check
                    int? cellFill = null;
                    try
                    {
                        if (cell.Interior.Color != null && !(cell.Interior.ColorIndex is int idx && idx == (int)Excel.Constants.xlNone))
                        {
                            cellFill = Convert.ToInt32(cell.Interior.Color);
                        }
                    }
                    catch { }

                    if (criteria.MatchFillColor)
                    {
                        if (criteria.FillColorRgb.HasValue && cellFill != criteria.FillColorRgb.Value)
                            isMatch = false;
                        else if (!criteria.FillColorRgb.HasValue && cellFill.HasValue)
                            isMatch = false;
                    }

                    // Font Color check
                    int? cellFontColor = null;
                    try
                    {
                        if (cell.Font.Color != null)
                        {
                            cellFontColor = Convert.ToInt32(cell.Font.Color);
                        }
                    }
                    catch { }

                    if (isMatch && criteria.MatchFontColor)
                    {
                        if (criteria.FontColorRgb.HasValue && cellFontColor != criteria.FontColorRgb.Value)
                            isMatch = false;
                    }

                    // Bold check
                    bool cellBold = false;
                    try { cellBold = cell.Font.Bold is bool b && b; } catch { }

                    if (isMatch && criteria.MatchBold)
                    {
                        if (criteria.IsBold.HasValue && cellBold != criteria.IsBold.Value)
                            isMatch = false;
                    }

                    // Italic check
                    bool cellItalic = false;
                    try { cellItalic = cell.Font.Italic is bool it && it; } catch { }

                    if (isMatch && criteria.MatchItalic)
                    {
                        if (criteria.IsItalic.HasValue && cellItalic != criteria.IsItalic.Value)
                            isMatch = false;
                    }

                    if (isMatch)
                    {
                        string addr = cell.get_Address();
                        string val = cell.Value2?.ToString() ?? "";
                        matches.Add(new FormatMatchItem(cell.Row, cell.Column, addr, sheetName, val, cellFill, cellFontColor, cellBold, cellItalic));
                    }
                }
            }
            catch { }
            return matches;
        }

        public CellFormatCriteria GetActiveCellFormatCriteria()
        {
            var criteria = new CellFormatCriteria();
            if (!HasActiveWorksheet) return criteria;

            try
            {
                var cell = _app.ActiveCell;
                if (cell != null)
                {
                    try
                    {
                        if (cell.Interior.Color != null && !(cell.Interior.ColorIndex is int idx && idx == (int)Excel.Constants.xlNone))
                        {
                            criteria.MatchFillColor = true;
                            criteria.FillColorRgb = Convert.ToInt32(cell.Interior.Color);
                        }
                    }
                    catch { }

                    try
                    {
                        if (cell.Font.Color != null)
                        {
                            criteria.MatchFontColor = true;
                            criteria.FontColorRgb = Convert.ToInt32(cell.Font.Color);
                        }
                    }
                    catch { }

                    try
                    {
                        if (cell.Font.Bold is bool b && b)
                        {
                            criteria.MatchBold = true;
                            criteria.IsBold = true;
                        }
                    }
                    catch { }

                    try
                    {
                        if (cell.Font.Italic is bool it && it)
                        {
                            criteria.MatchItalic = true;
                            criteria.IsItalic = true;
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return criteria;
        }

        public IReadOnlyList<ValidationRuleInfo> GetWorksheetValidationRules(bool selectionOnly = false)
        {
            var list = new List<ValidationRuleInfo>();
            if (!HasActiveWorksheet) return list;

            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                Excel.Range baseRange = selectionOnly ? (_app.Selection as Excel.Range ?? ws.UsedRange) : ws.UsedRange;
                if (baseRange == null) return list;

                Excel.Range? valRange = null;
                try
                {
                    valRange = baseRange.SpecialCells(Excel.XlCellType.xlCellTypeAllValidation);
                }
                catch (COMException)
                {
                    valRange = null;
                }

                if (valRange != null)
                {
                    string sheetName = ws.Name;
                    int count = 0;

                    foreach (Excel.Range cell in valRange.Cells)
                    {
                        count++;
                        if (count > 2000) break; // Performance safeguard

                        string typeName = "Custom";
                        string opName = "";
                        string f1 = "";
                        string f2 = "";
                        string errMsg = "";
                        bool? isValid = null;

                        try
                        {
                            var v = cell.Validation;
                            int typeVal = v.Type;
                            typeName = typeVal switch
                            {
                                0 => "Any Value",
                                1 => "Whole Number",
                                2 => "Decimal",
                                3 => "List",
                                4 => "Date",
                                5 => "Time",
                                6 => "Text Length",
                                7 => "Custom",
                                _ => "Other"
                            };

                            try { opName = v.Operator.ToString(); } catch { }
                            try { f1 = v.Formula1 ?? ""; } catch { }
                            try { f2 = v.Formula2 ?? ""; } catch { }
                            try { errMsg = v.ErrorMessage ?? ""; } catch { }
                            try { isValid = v.Value; } catch { }
                        }
                        catch { }

                        string valStr = cell.Value2?.ToString() ?? "";
                        list.Add(new ValidationRuleInfo(sheetName, cell.get_Address(), cell.Row, cell.Column, typeName, opName, f1, f2, valStr, errMsg, isValid));
                    }
                }
            }
            catch { }
            return list;
        }

        public void SelectCellsByAddresses(IEnumerable<string> addresses)
        {
            if (!HasActiveWorksheet || addresses == null) return;

            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                var list = new List<string>(addresses);
                if (list.Count == 0) return;

                // Group addresses into chunks to respect Excel's 255-character Range string limit
                Excel.Range? unionRange = null;
                var currentChunk = new List<string>();
                int currentLen = 0;

                foreach (var addr in list)
                {
                    if (currentLen + addr.Length + 1 > 240)
                    {
                        var chunkRange = ws.Range[string.Join(",", currentChunk)];
                        unionRange = unionRange == null ? chunkRange : _app.Union(unionRange, chunkRange);
                        currentChunk.Clear();
                        currentLen = 0;
                    }
                    currentChunk.Add(addr);
                    currentLen += addr.Length + 1;
                }

                if (currentChunk.Count > 0)
                {
                    var chunkRange = ws.Range[string.Join(",", currentChunk)];
                    unionRange = unionRange == null ? chunkRange : _app.Union(unionRange, chunkRange);
                }

                unionRange?.Select();
            }
            catch { }
        }

        public void RemoveValidationFromRange(IEnumerable<string> addresses)
        {
            if (!HasActiveWorksheet || addresses == null) return;

            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                using (new ExcelStateScope(_app))
                {
                    foreach (var addr in addresses)
                    {
                        try
                        {
                            ws.Range[addr].Validation.Delete();
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to remove validation: {ex.Message}", "Validation Error", isError: true);
            }
        }

        public int ClearValidationRules(ValidationClearScope scope)
        {
            if (!HasActiveWorkbook) return 0;

            try
            {
                using (new ExcelStateScope(_app))
                {
                    switch (scope)
                    {
                        case ValidationClearScope.Selection:
                            if (_app.Selection is Excel.Range sel)
                            {
                                sel.Validation.Delete();
                                return 1;
                            }
                            break;

                        case ValidationClearScope.ActiveWorksheet:
                            if (_app.ActiveSheet is Excel.Worksheet ws)
                            {
                                ws.Cells.Validation.Delete();
                                return 1;
                            }
                            break;

                        case ValidationClearScope.AllWorksheets:
                            int cleared = 0;
                            var wb = _app.ActiveWorkbook;
                            foreach (Excel.Worksheet sheet in wb.Worksheets)
                            {
                                try
                                {
                                    sheet.Cells.Validation.Delete();
                                    cleared++;
                                }
                                catch { }
                            }
                            return cleared;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to clear validation rules: {ex.Message}", "Validation Clear Error", isError: true);
            }

            return 0;
        }

        public bool CreateDropdownList(string listFormulaOrItems, bool isFormula, bool allowBlank, bool showDropdown, string? errorTitle, string? errorMessage)
        {
            if (!HasActiveWorksheet) return false;
            if (string.IsNullOrWhiteSpace(listFormulaOrItems)) return false;

            try
            {
                if (!(_app.Selection is Excel.Range range)) return false;
                if (IsCurrentSheetProtected)
                {
                    ShowMessage("The active worksheet is protected. Cannot modify validation rules.", "Sheet Protected", isError: true);
                    return false;
                }

                using (new ExcelStateScope(_app))
                {
                    Excel.Validation val = range.Validation;
                    val.Delete();
                    val.Add(
                        Excel.XlDVType.xlValidateList,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlBetween,
                        listFormulaOrItems);

                    val.IgnoreBlank = allowBlank;
                    val.InCellDropdown = showDropdown;

                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        val.ShowError = true;
                        val.ErrorTitle = errorTitle ?? "Invalid Entry";
                        val.ErrorMessage = errorMessage;
                    }
                    else
                    {
                        val.ShowError = false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to create dropdown list: {ex.Message}", "Dropdown List Error", isError: true);
                return false;
            }
        }

        public bool AdjustCellDimensions(double? rowHeightPoints, double? colWidthPoints, bool autoFitRows, bool autoFitCols, bool entireSheet)
        {
            if (!HasActiveWorksheet) return false;

            try
            {
                var ws = (Excel.Worksheet)_app.ActiveSheet;
                Excel.Range target = entireSheet ? ws.Cells : (_app.Selection as Excel.Range ?? ws.UsedRange);

                using (new ExcelStateScope(_app))
                {
                    if (autoFitRows)
                    {
                        target.Rows.AutoFit();
                    }
                    else if (rowHeightPoints.HasValue && rowHeightPoints.Value > 0)
                    {
                        target.RowHeight = Math.Min(409.5, Math.Max(1.0, rowHeightPoints.Value));
                    }

                    if (autoFitCols)
                    {
                        target.Columns.AutoFit();
                    }
                    else if (colWidthPoints.HasValue && colWidthPoints.Value > 0)
                    {
                        // In Excel COM, ColumnWidth is in characters (~1 char ~= 5.3 points for 11pt Calibri default)
                        // Standard conversion: characters = points / 5.25
                        double chars = colWidthPoints.Value / 5.25;
                        target.ColumnWidth = Math.Min(255.0, Math.Max(0.5, chars));
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to adjust cell dimensions: {ex.Message}", "Dimension Error", isError: true);
                return false;
            }
        }

        public OperationResult InsertSequenceNumbers(SequenceNumberDefinition sequence, SequenceFillOptions options)
        {
            if (sequence == null) throw new ArgumentNullException(nameof(sequence));
            options ??= new SequenceFillOptions();

            if (!HasActiveWorkbook || !(_app.ActiveSheet is Excel.Worksheet sheet))
            {
                ShowMessage("No active worksheet is currently open.", "Insert Sequence Number", isError: true);
                return OperationResult.Failed("No active worksheet found.");
            }

            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before modifying cell values.", "Sheet Protected", isError: true);
                return OperationResult.Failed("The active worksheet is protected.");
            }

            if (!(_app.Selection is Excel.Range selRange))
            {
                ShowMessage("Please select a range of cells where you want to insert sequential numbers.", "No Selection", isError: true);
                return OperationResult.Failed("No cells are selected.");
            }

            options.WorksheetName = sheet.Name;
            options.ReferenceDate = DateTime.Now;

            try
            {
                // FAST PATH: Single contiguous area without merged cells, non-filtered, non-grouped
                bool isMerged = selRange.MergeCells is true || selRange.MergeCells is null || (selRange.MergeCells is object o && Convert.IsDBNull(o));
                if (selRange.Areas.Count == 1 && !isMerged && !options.OnlyFillBlankCells && !options.SkipHiddenCells && !options.RestartOnValueChange && options.StepRows <= 1)
                {
                    object?[,] beforeValues = NormalizeTo2DArray(selRange.Value2);
                    object?[,] newMatrix = SequenceNumberEngine.FormatMatrix(
                        beforeValues,
                        sequence,
                        options,
                        out int updated,
                        out int skipped,
                        out long nextNum);

                    RecordBeforeRangeChange(selRange, newMatrix, isFormula: false, customDesc: "Insert Sequence Numbers");

                    using (new ExcelStateScope(_app))
                    {
                        selRange.Value2 = newMatrix;
                        if (sequence.OutputAsNumeric && sequence.Type == SequenceType.Numeric && sequence.Digits > 1)
                        {
                            try { selRange.NumberFormat = new string('0', sequence.Digits); } catch { }
                        }
                    }

                    sequence.CurrentNumber = nextNum;
                    return OperationResult.Succeeded($"Successfully inserted {updated} sequence numbers.", updated, skipped);
                }

                // ADVANCED PATH: Merged cells, non-contiguous multi-area selections, blank filtering, serpentine traversal, or grouping
                var targetCells = new List<Excel.Range>();
                var seenMergeAreas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                void ProcessCell(Excel.Range area, int r, int c)
                {
                    if (options.StepRows > 1 && ((r - 1) % options.StepRows) != 0)
                        return;

                    Excel.Range cell = area.Cells[r, c];
                    if (options.SkipHiddenCells && ((bool)cell.EntireRow.Hidden || (bool)cell.EntireColumn.Hidden))
                        return;

                    if (cell.MergeCells is true)
                    {
                        Excel.Range mergeArea = cell.MergeArea;
                        string mergeAddr = mergeArea.get_Address();
                        if (seenMergeAreas.Contains(mergeAddr)) return;
                        seenMergeAreas.Add(mergeAddr);
                        cell = (Excel.Range)mergeArea.Cells[1, 1];
                    }

                    if (options.OnlyFillBlankCells)
                    {
                        object? val = cell.Value2;
                        if (val != null && !string.IsNullOrWhiteSpace(val.ToString())) return;
                    }

                    targetCells.Add(cell);
                }

                foreach (Excel.Range area in selRange.Areas)
                {
                    int rCount = area.Rows.Count;
                    int cCount = area.Columns.Count;

                    switch (options.FillOrder)
                    {
                        case SequenceFillOrder.Vertical:
                            for (int c = 1; c <= cCount; c++)
                                for (int r = 1; r <= rCount; r++)
                                    ProcessCell(area, r, c);
                            break;

                        case SequenceFillOrder.Horizontal:
                            for (int r = 1; r <= rCount; r++)
                                for (int c = 1; c <= cCount; c++)
                                    ProcessCell(area, r, c);
                            break;

                        case SequenceFillOrder.SerpentineVertical:
                            for (int c = 1; c <= cCount; c++)
                            {
                                if (c % 2 != 0)
                                {
                                    for (int r = 1; r <= rCount; r++)
                                        ProcessCell(area, r, c);
                                }
                                else
                                {
                                    for (int r = rCount; r >= 1; r--)
                                        ProcessCell(area, r, c);
                                }
                            }
                            break;

                        case SequenceFillOrder.SerpentineHorizontal:
                            for (int r = 1; r <= rCount; r++)
                            {
                                if (r % 2 != 0)
                                {
                                    for (int c = 1; c <= cCount; c++)
                                        ProcessCell(area, r, c);
                                }
                                else
                                {
                                    for (int c = cCount; c >= 1; c--)
                                        ProcessCell(area, r, c);
                                }
                            }
                            break;
                    }
                }

                if (targetCells.Count == 0)
                {
                    return OperationResult.Failed("No eligible cells found in the selection to insert sequential numbers.");
                }

                using (new ExcelStateScope(_app))
                {
                    long currentVal = options.ContinueFromLast ? sequence.CurrentNumber : sequence.StartNumber;
                    int rep = 0;
                    int updated = 0;

                    if (options.RandomizeOrder && !options.RestartOnValueChange)
                    {
                        var valuesList = new List<object>();
                        long genVal = currentVal;
                        int genRep = 0;

                        for (int i = 0; i < targetCells.Count; i++)
                        {
                            string formatted = sequence.FormatValue(genVal, options.ReferenceDate, options.WorksheetName);
                            if (sequence.OutputAsNumeric && sequence.Type == SequenceType.Numeric && long.TryParse(formatted, out long num))
                            {
                                valuesList.Add(num);
                            }
                            else
                            {
                                valuesList.Add(formatted);
                            }

                            genRep++;
                            if (genRep >= sequence.RepeatCount)
                            {
                                genRep = 0;
                                genVal += sequence.Increment;
                                if (sequence.EndNumber.HasValue)
                                {
                                    if (sequence.Increment > 0 && genVal > sequence.EndNumber.Value) genVal = sequence.StartNumber;
                                    else if (sequence.Increment < 0 && genVal < sequence.EndNumber.Value) genVal = sequence.StartNumber;
                                }
                            }
                        }

                        var rng = new Random();
                        for (int i = valuesList.Count - 1; i > 0; i--)
                        {
                            int j = rng.Next(i + 1);
                            var tmp = valuesList[i];
                            valuesList[i] = valuesList[j];
                            valuesList[j] = tmp;
                        }

                        for (int i = 0; i < targetCells.Count; i++)
                        {
                            var cell = targetCells[i];
                            cell.Value2 = valuesList[i];
                            if (sequence.OutputAsNumeric && sequence.Type == SequenceType.Numeric && sequence.Digits > 1)
                            {
                                try { cell.NumberFormat = new string('0', sequence.Digits); } catch { }
                            }
                            updated++;
                        }

                        sequence.CurrentNumber = genVal;
                    }
                    else
                    {
                        string? prevKeyStr = null;
                        bool hasPrevKey = false;

                        foreach (var cell in targetCells)
                        {
                            if (options.RestartOnValueChange)
                            {
                                try
                                {
                                    Excel.Range keyCell = cell.Offset[0, options.RestartKeyColumnOffset];
                                    object? keyVal = keyCell?.Value2;
                                    string keyStr = keyVal?.ToString() ?? string.Empty;
                                    if (hasPrevKey && !string.Equals(keyStr, prevKeyStr, StringComparison.OrdinalIgnoreCase))
                                    {
                                        currentVal = sequence.StartNumber;
                                        rep = 0;
                                    }
                                    prevKeyStr = keyStr;
                                    hasPrevKey = true;
                                }
                                catch { }
                            }

                            string formatted = sequence.FormatValue(currentVal, options.ReferenceDate, options.WorksheetName);
                            if (sequence.OutputAsNumeric && sequence.Type == SequenceType.Numeric && long.TryParse(formatted, out long num))
                            {
                                cell.Value2 = num;
                                if (sequence.Digits > 1)
                                {
                                    try { cell.NumberFormat = new string('0', sequence.Digits); } catch { }
                                }
                            }
                            else
                            {
                                cell.Value2 = formatted;
                            }

                            updated++;

                            rep++;
                            if (rep >= sequence.RepeatCount)
                            {
                                rep = 0;
                                currentVal += sequence.Increment;
                                if (sequence.EndNumber.HasValue)
                                {
                                    if (sequence.Increment > 0 && currentVal > sequence.EndNumber.Value) currentVal = sequence.StartNumber;
                                    else if (sequence.Increment < 0 && currentVal < sequence.EndNumber.Value) currentVal = sequence.StartNumber;
                                }
                            }
                        }

                        sequence.CurrentNumber = currentVal;
                    }

                    return OperationResult.Succeeded($"Successfully inserted {updated} sequence numbers.", updated);
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error inserting sequence numbers: {ex.Message}", "Insert Sequence Error", isError: true);
                return OperationResult.Failed(ex.Message);
            }
        }

        public OperationResult TransformRange(TransformRangeOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            var (isValid, errorMessage) = options.Validate();
            if (!isValid)
            {
                ShowMessage(errorMessage ?? "Invalid transformation options.", "Transform Range", isError: true);
                return OperationResult.Failed(errorMessage ?? "Invalid transformation options.");
            }

            if (IsCurrentSheetProtected)
            {
                ShowMessage("The active worksheet is protected. Please unprotect the sheet before transforming.", "Sheet Protected", isError: true);
                return OperationResult.Failed("Worksheet is protected.");
            }

            var matrix = GetSelectedValues();
            if (matrix == null || matrix.Length == 0)
            {
                ShowMessage("No cells selected to transform.", "Transform Range", isError: true);
                return OperationResult.Failed("No cells selected.");
            }

            try
            {
                var transformed = RangeTransformer.Transform(matrix, options);
                int outRows = transformed.GetLength(0);
                int outCols = transformed.GetLength(1);

                if (outRows == 0 || outCols == 0)
                {
                    ShowMessage("Transformation yielded no records.", "Transform Range", isError: false);
                    return OperationResult.Succeeded("No records transformed.", 0);
                }

                if (options.DestinationMode == TransformDestinationMode.NewSheet)
                {
                    string sheetName = string.IsNullOrWhiteSpace(options.NewSheetName) ? "Transformed" : options.NewSheetName.Trim();
                    bool created = CreateWorksheetWithData(sheetName, transformed);
                    if (!created)
                    {
                        ShowMessage($"Could not create sheet '{sheetName}'. A sheet with this name may already exist.", "Transform Range", isError: true);
                        return OperationResult.Failed($"Worksheet '{sheetName}' could not be created.");
                    }
                }
                else if (options.DestinationMode == TransformDestinationMode.TargetCell)
                {
                    string targetAddress = options.TargetAddress?.Trim() ?? "$A$1";
                    SetRangeValuesByAddress(targetAddress, transformed);
                }
                else // InPlace
                {
                    int startRow = SelectionStartRow;
                    int startCol = SelectionStartColumn;
                    ClearSelection();
                    SetRangeValues(startRow, startCol, transformed);
                }

                int totalCells = outRows * outCols;
                return OperationResult.Succeeded($"Successfully transformed {outRows} × {outCols} range ({totalCells} cells).", totalCells);
            }
            catch (Exception ex)
            {
                ShowMessage($"Error transforming range: {ex.Message}", "Transform Range Error", isError: true);
                return OperationResult.Failed(ex.Message);
            }
        }

        private static object?[,] NormalizeTo2DArray(object? raw)
        {
            if (raw is object[,] matrix)
            {
                int rStart = matrix.GetLowerBound(0);
                int rEnd = matrix.GetUpperBound(0);
                int cStart = matrix.GetLowerBound(1);
                int cEnd = matrix.GetUpperBound(1);

                int rows = rEnd - rStart + 1;
                int cols = cEnd - cStart + 1;
                var res = new object?[rows, cols];
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        res[r, c] = matrix[rStart + r, cStart + c];
                    }
                }
                return res;
            }

            var single = new object?[1, 1];
            single[0, 0] = raw;
            return single;
        }

        public const int MaxUndoCellCount = 500_000;

        private void RecordBeforeRangeChange(Excel.Range range, object?[,] newValues, bool isFormula = false, string? customDesc = null)
        {
            if (range == null || newValues == null) return;
            try
            {
                long totalCells = (long)range.Rows.Count * range.Columns.Count;
                if (totalCells > MaxUndoCellCount)
                {
                    // Safeguard: Skip recording massive ranges to protect memory on 32-bit Excel 2016
                    return;
                }

                var sheet = (Excel.Worksheet)range.Worksheet;
                var wb = (Excel.Workbook)sheet.Parent;
                string wbName = wb?.Name ?? string.Empty;
                string? wbPath = null;
                try { wbPath = wb?.FullName; } catch { }

                object?[,] beforeValues = NormalizeTo2DArray(range.Value2);
                object?[,]? beforeFormulas = null;

                try
                {
                    object hasFormObj = range.HasFormula;
                    // In Excel COM:
                    // bool false = no formulas in range
                    // bool true = all cells have formulas
                    // DBNull.Value = mixed range with some formulas and some constants
                    bool hasFormulas = !(hasFormObj is bool b && !b);
                    if (hasFormulas)
                    {
                        var formulaMatrix = NormalizeTo2DArray(range.Formula);
                        bool containsRealFormula = false;
                        for (int r = 0; r < formulaMatrix.GetLength(0); r++)
                        {
                            for (int c = 0; c < formulaMatrix.GetLength(1); c++)
                            {
                                if (formulaMatrix[r, c] is string s && s.StartsWith("="))
                                {
                                    containsRealFormula = true;
                                    break;
                                }
                            }
                            if (containsRealFormula) break;
                        }

                        if (containsRealFormula)
                        {
                            beforeFormulas = formulaMatrix;
                        }
                    }
                }
                catch { }

                object?[,]? afterFormulas = isFormula ? newValues : null;
                object?[,] afterValues = isFormula ? beforeValues : newValues;

                string desc = customDesc ?? (isFormula ? "Formula Update" : "Cell Value Update");
                _undoManager.RecordAction(
                    description: desc,
                    workbookName: wbName,
                    workbookPath: wbPath,
                    sheetName: sheet.Name,
                    rangeAddress: range.get_Address(),
                    startRow: range.Row,
                    startColumn: range.Column,
                    beforeValues: beforeValues,
                    afterValues: afterValues,
                    beforeFormulas: beforeFormulas,
                    afterFormulas: afterFormulas);
            }
            catch
            {
                // Non-blocking undo tracking
            }
        }

        public bool UndoLastAction()
        {
            if (!HasActiveWorkbook) return false;

            if (!_undoManager.CanUndo)
            {
                try
                {
                    _app.Undo();
                    return true;
                }
                catch
                {
                    ShowMessage("There are no suite actions to undo.", "Undo", isError: false);
                    return false;
                }
            }

            var step = _undoManager.PopUndo();
            if (step == null) return false;

            try
            {
                Excel.Workbook? targetWb = null;
                if (!string.IsNullOrEmpty(step.WorkbookName))
                {
                    foreach (Excel.Workbook w in _app.Workbooks)
                    {
                        if (string.Equals(w.Name, step.WorkbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWb = w;
                            break;
                        }
                    }
                }
                targetWb ??= _app.ActiveWorkbook;

                if (targetWb == null)
                {
                    ShowMessage($"Cannot undo '{step.Description}': Target workbook '{step.WorkbookName}' is closed or unavailable.", "Undo Error", isError: true);
                    return false;
                }

                Excel.Worksheet? targetSheet = null;
                foreach (Excel.Worksheet s in targetWb.Worksheets)
                {
                    if (string.Equals(s.Name, step.SheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetSheet = s;
                        break;
                    }
                }

                if (targetSheet == null)
                {
                    ShowMessage($"Cannot undo '{step.Description}': Worksheet '{step.SheetName}' was not found in workbook '{targetWb.Name}'.", "Undo Error", isError: true);
                    return false;
                }

                if (targetSheet.ProtectContents)
                {
                    ShowMessage($"Cannot undo '{step.Description}': Worksheet '{step.SheetName}' is protected. Please unprotect the sheet first.", "Sheet Protected", isError: true);
                    return false;
                }

                using (new ExcelStateScope(_app))
                {
                    try { targetSheet.Activate(); } catch { }

                    Excel.Range targetRange = targetSheet.Range[step.RangeAddress];

                    _undoManager.PushRedo(step);

                    if (step.BeforeFormulas != null)
                    {
                        targetRange.Formula = step.BeforeFormulas;
                    }
                    else
                    {
                        targetRange.Value2 = step.BeforeValues;
                    }
                }

                try
                {
                    _app.StatusBar = $"Suite Undo: Undid '{step.Description}' on [{targetWb.Name}]{step.SheetName}!{step.RangeAddress}";
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Error performing undo: {ex.Message}", "Undo Error", isError: true);
                return false;
            }
        }

        public bool RedoLastAction()
        {
            if (!HasActiveWorkbook) return false;

            if (!_undoManager.CanRedo)
            {
                ShowMessage("There are no suite actions to redo.", "Redo", isError: false);
                return false;
            }

            var step = _undoManager.PopRedo();
            if (step == null) return false;

            try
            {
                Excel.Workbook? targetWb = null;
                if (!string.IsNullOrEmpty(step.WorkbookName))
                {
                    foreach (Excel.Workbook w in _app.Workbooks)
                    {
                        if (string.Equals(w.Name, step.WorkbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWb = w;
                            break;
                        }
                    }
                }
                targetWb ??= _app.ActiveWorkbook;

                if (targetWb == null)
                {
                    ShowMessage($"Cannot redo '{step.Description}': Target workbook '{step.WorkbookName}' is closed or unavailable.", "Redo Error", isError: true);
                    return false;
                }

                Excel.Worksheet? targetSheet = null;
                foreach (Excel.Worksheet s in targetWb.Worksheets)
                {
                    if (string.Equals(s.Name, step.SheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetSheet = s;
                        break;
                    }
                }

                if (targetSheet == null)
                {
                    ShowMessage($"Cannot redo '{step.Description}': Worksheet '{step.SheetName}' was not found in workbook '{targetWb.Name}'.", "Redo Error", isError: true);
                    return false;
                }

                if (targetSheet.ProtectContents)
                {
                    ShowMessage($"Cannot redo '{step.Description}': Worksheet '{step.SheetName}' is protected. Please unprotect the sheet first.", "Sheet Protected", isError: true);
                    return false;
                }

                using (new ExcelStateScope(_app))
                {
                    try { targetSheet.Activate(); } catch { }

                    Excel.Range targetRange = targetSheet.Range[step.RangeAddress];

                    _undoManager.PushUndo(step);

                    if (step.AfterFormulas != null)
                    {
                        targetRange.Formula = step.AfterFormulas;
                    }
                    else
                    {
                        targetRange.Value2 = step.AfterValues;
                    }
                }

                try
                {
                    _app.StatusBar = $"Suite Redo: Redid '{step.Description}' on [{targetWb.Name}]{step.SheetName}!{step.RangeAddress}";
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                ShowMessage($"Error performing redo: {ex.Message}", "Redo Error", isError: true);
                return false;
            }
        }

        public IReadOnlyList<string> GetInstalledPrinterNames()
        {
            var list = new List<string>();
            try
            {
                foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    list.Add(printer);
                }
            }
            catch { }
            return list;
        }

        public string? GetDefaultPrinterName()
        {
            try
            {
                var settings = new System.Drawing.Printing.PrinterSettings();
                return settings.PrinterName;
            }
            catch
            {
                return null;
            }
        }

        public PrintWorkbooksResult BatchPrintWorkbooks(
            PrintWorkbooksConfig config,
            IProgress<(int current, int total, string currentFile)>? progress = null,
            CancellationToken cancellationToken = default)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var result = new PrintWorkbooksResult();

            var (isValid, errorMsg) = PrintWorkbooksEngine.ValidateConfig(config);
            if (!isValid)
            {
                result.Errors.Add(errorMsg ?? "Invalid configuration.");
                return result;
            }

            var targets = config.SelectedFiles.Where(f => f.IsSelected).ToList();
            result.TotalWorkbooks = targets.Count;

            string? originalPrinter = null;
            bool originalScreenUpdating = _app.ScreenUpdating;
            bool originalDisplayAlerts = _app.DisplayAlerts;
            bool originalEnableEvents = _app.EnableEvents;

            try
            {
                if (!string.IsNullOrWhiteSpace(config.PrinterName))
                {
                    try
                    {
                        originalPrinter = _app.ActivePrinter;
                        _app.ActivePrinter = config.PrinterName;
                    }
                    catch { }
                }

                if (config.QuietMode)
                {
                    _app.ScreenUpdating = false;
                    _app.DisplayAlerts = false;
                }

                if (config.DisableMacros)
                {
                    _app.EnableEvents = false;
                }

                int currentIdx = 0;
                foreach (var item in targets)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        item.Status = PrintItemStatus.Skipped;
                        item.ErrorMessage = "Operation cancelled by user.";
                        result.SkippedWorkbooks++;
                        continue;
                    }

                    currentIdx++;
                    progress?.Report((currentIdx, targets.Count, item.FileName));

                    item.Status = PrintItemStatus.Printing;
                    Excel.Workbook? wb = null;
                    try
                    {
                        int updateLinks = config.DoNotUpdateLinks ? 0 : 3;
                        wb = _app.Workbooks.Open(item.FilePath, UpdateLinks: updateLinks, ReadOnly: true);

                        int sheetsPrintedInWb = 0;
                        var (fromPage, toPage) = PrintWorkbooksEngine.ResolvePageRange(config.PageRangeMode, config.FromPage, config.ToPage);

                        foreach (object sheetObj in wb.Sheets)
                        {
                            if (cancellationToken.IsCancellationRequested) break;

                            string sheetName = string.Empty;
                            bool isChart = sheetObj is Excel.Chart;

                            if (sheetObj is Excel.Worksheet ws)
                            {
                                sheetName = ws.Name;
                            }
                            else if (sheetObj is Excel.Chart chart)
                            {
                                sheetName = chart.Name;
                            }

                            if (PrintWorkbooksEngine.ShouldPrintSheet(sheetName, isChart, config.SheetScope, config.SpecificSheetNames))
                            {
                                try
                                {
                                    if (fromPage.HasValue && toPage.HasValue)
                                    {
                                        if (sheetObj is Excel.Worksheet wsPrint)
                                        {
                                            wsPrint.PrintOut(From: fromPage.Value, To: toPage.Value, Copies: config.Copies);
                                        }
                                        else if (sheetObj is Excel.Chart chartPrint)
                                        {
                                            chartPrint.PrintOut(From: fromPage.Value, To: toPage.Value, Copies: config.Copies);
                                        }
                                    }
                                    else
                                    {
                                        if (sheetObj is Excel.Worksheet wsPrint)
                                        {
                                            wsPrint.PrintOut(Copies: config.Copies);
                                        }
                                        else if (sheetObj is Excel.Chart chartPrint)
                                        {
                                            chartPrint.PrintOut(Copies: config.Copies);
                                        }
                                    }
                                    sheetsPrintedInWb++;
                                }
                                catch (Exception printEx)
                                {
                                    item.ErrorMessage = $"Sheet '{sheetName}': {printEx.Message}";
                                }
                            }
                        }

                        item.SheetsPrintedCount = sheetsPrintedInWb;
                        item.Status = PrintItemStatus.Completed;
                        result.PrintedWorkbooks++;
                        result.TotalSheetsPrinted += sheetsPrintedInWb;
                    }
                    catch (Exception fileEx)
                    {
                        item.Status = PrintItemStatus.Failed;
                        item.ErrorMessage = fileEx.Message;
                        result.FailedWorkbooks++;
                        result.Errors.Add($"{item.FileName}: {fileEx.Message}");
                    }
                    finally
                    {
                        if (wb != null)
                        {
                            try
                            {
                                wb.Close(SaveChanges: false);
                                Marshal.ReleaseComObject(wb);
                            }
                            catch { }
                        }
                    }
                }
            }
            finally
            {
                if (originalPrinter != null)
                {
                    try { _app.ActivePrinter = originalPrinter; } catch { }
                }
                _app.ScreenUpdating = originalScreenUpdating;
                _app.DisplayAlerts = originalDisplayAlerts;
                _app.EnableEvents = originalEnableEvents;

                sw.Stop();
                result.Duration = sw.Elapsed;
                result.ProcessedItems = targets;
            }

            if (config.GenerateSummaryReport && result.ProcessedItems.Count > 0)
            {
                try
                {
                    var matrix = PrintWorkbooksEngine.GenerateSummaryMatrix(result, result.ProcessedItems);
                    CreateWorksheetWithData("Print_Summary", matrix, activate: true);
                }
                catch { }
            }

            return result;
        }

        public EmptyRowsScanResult ScanEmptyRows(DeleteEmptyRowsOptions options)
        {
            options ??= new DeleteEmptyRowsOptions();

            if (!HasActiveWorkbook || !HasActiveWorksheet)
            {
                return new EmptyRowsScanResult { WorksheetName = ActiveWorksheetName ?? string.Empty };
            }

            try
            {
                Excel.Worksheet? ws = _app.ActiveSheet as Excel.Worksheet;
                if (ws == null) return new EmptyRowsScanResult();

                if (options.Scope == DeleteEmptyRowsScope.CurrentSelection && _app.Selection is Excel.Range sel)
                {
                    int startRow = sel.Row;
                    int startCol = sel.Column;
                    object?[,] matrix = NormalizeTo2DArray(sel.Value2);
                    return DeleteEmptyRowsEngine.ScanWorksheetMatrix(
                        ws.Name,
                        matrix,
                        startRow,
                        startCol,
                        options.TreatWhitespaceAsEmpty,
                        options.TreatEmptyFormulaAsEmpty);
                }
                else
                {
                    Excel.Range used = ws.UsedRange;
                    if (used == null)
                    {
                        return new EmptyRowsScanResult { WorksheetName = ws.Name };
                    }

                    int startRow = used.Row;
                    int startCol = used.Column;
                    object?[,] matrix = NormalizeTo2DArray(used.Value2);

                    return DeleteEmptyRowsEngine.ScanWorksheetMatrix(
                        ws.Name,
                        matrix,
                        startRow,
                        startCol,
                        options.TreatWhitespaceAsEmpty,
                        options.TreatEmptyFormulaAsEmpty);
                }
            }
            catch
            {
                return new EmptyRowsScanResult
                {
                    WorksheetName = ActiveWorksheetName ?? string.Empty,
                    TotalRowsScanned = 0
                };
            }
        }

        public DeleteEmptyRowsResult DeleteAllEmptyRows(DeleteEmptyRowsOptions options)
        {
            options ??= new DeleteEmptyRowsOptions();

            if (!HasActiveWorkbook)
            {
                return DeleteEmptyRowsResult.Failed("No active workbook is open in Excel.");
            }

            var targetSheets = new List<Excel.Worksheet>();

            if (options.Scope == DeleteEmptyRowsScope.AllWorksheetsInWorkbook)
            {
                foreach (Excel.Worksheet ws in _app.ActiveWorkbook.Worksheets)
                {
                    targetSheets.Add(ws);
                }
            }
            else if (options.Scope == DeleteEmptyRowsScope.SelectedWorksheets && options.TargetWorksheets != null && options.TargetWorksheets.Count > 0)
            {
                foreach (Excel.Worksheet ws in _app.ActiveWorkbook.Worksheets)
                {
                    if (options.TargetWorksheets.Contains(ws.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        targetSheets.Add(ws);
                    }
                }
            }
            else
            {
                Excel.Worksheet? activeWs = _app.ActiveSheet as Excel.Worksheet;
                if (activeWs == null)
                {
                    return DeleteEmptyRowsResult.Failed("No active worksheet found.");
                }
                targetSheets.Add(activeWs);
            }

            var protectedSheets = targetSheets.Where(s => s.ProtectContents).Select(s => s.Name).ToList();
            if (protectedSheets.Count > 0)
            {
                return DeleteEmptyRowsResult.Failed($"The following worksheet(s) are protected: {string.Join(", ", protectedSheets)}. Please unprotect them before deleting rows.");
            }

            int grandTotalDeleted = 0;
            var perSheetDeleted = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            using (new ExcelStateScope(_app))
            {
                foreach (var ws in targetSheets)
                {
                    try
                    {
                        Excel.Range? targetRange;
                        if (options.Scope == DeleteEmptyRowsScope.CurrentSelection && targetSheets.Count == 1 && _app.Selection is Excel.Range sel)
                        {
                            targetRange = sel;
                        }
                        else
                        {
                            targetRange = ws.UsedRange;
                        }

                        if (targetRange == null) continue;

                        int startRow = targetRange.Row;
                        int startCol = targetRange.Column;
                        object?[,] matrix = NormalizeTo2DArray(targetRange.Value2);

                        var emptyRowIndices = DeleteEmptyRowsEngine.ScanEmptyRows(
                            matrix,
                            startRow,
                            options.TreatWhitespaceAsEmpty,
                            options.TreatEmptyFormulaAsEmpty);

                        if (emptyRowIndices.Count == 0) continue;

                        if (targetSheets.Count == 1)
                        {
                            try
                            {
                                RecordBeforeRangeChange(targetRange, matrix, isFormula: false, customDesc: "Delete All Empty Rows");
                            }
                            catch { }
                        }

                        var contiguousRanges = DeleteEmptyRowsEngine.GroupContiguousRanges(emptyRowIndices);

                        if (options.Scope == DeleteEmptyRowsScope.CurrentSelection && targetRange.Columns.Count < ws.Columns.Count)
                        {
                            var reversedRanges = contiguousRanges.OrderByDescending(r => r.startRow).ToList();
                            foreach (var (rStart, rEnd) in reversedRanges)
                            {
                                int relStart = rStart - startRow + 1;
                                int relEnd = rEnd - startRow + 1;
                                int rowsInBlock = relEnd - relStart + 1;

                                Excel.Range block = targetRange.Rows[relStart];
                                if (rowsInBlock > 1)
                                {
                                    Excel.Range endBlock = targetRange.Rows[relEnd];
                                    block = ws.Range[block, endBlock];
                                }
                                block.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                            }
                        }
                        else
                        {
                            var chunks = DeleteEmptyRowsEngine.BuildChunkedAddresses(contiguousRanges, maxCharsPerChunk: 240, bottomToTop: true);

                            foreach (var chunk in chunks)
                            {
                                try
                                {
                                    Excel.Range rowsToDelete = ws.Range[chunk];
                                    rowsToDelete.EntireRow.Delete();
                                }
                                catch
                                {
                                    var splitParts = chunk.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var part in splitParts)
                                    {
                                        try
                                        {
                                            ws.Range[part.Trim()].EntireRow.Delete();
                                        }
                                        catch { }
                                    }
                                }
                            }
                        }

                        int sheetDeleted = emptyRowIndices.Count;
                        grandTotalDeleted += sheetDeleted;
                        perSheetDeleted[ws.Name] = sheetDeleted;
                    }
                    catch (Exception ex)
                    {
                        return DeleteEmptyRowsResult.Failed($"Error deleting rows on sheet '{ws.Name}': {ex.Message}");
                    }
                }
            }

            string successMsg = grandTotalDeleted == 0
                ? "No empty rows were found to delete."
                : (targetSheets.Count > 1
                    ? $"Successfully deleted {grandTotalDeleted} empty row(s) across {perSheetDeleted.Count} worksheet(s)."
                    : $"Successfully deleted {grandTotalDeleted} empty row(s) from '{targetSheets[0].Name}'.");

            return DeleteEmptyRowsResult.Succeeded(grandTotalDeleted, successMsg, perSheetDeleted);
        }

        public BatchInsertCheckboxResult InsertCheckboxes(BatchInsertCheckboxOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            if (!HasActiveWorkbook)
            {
                return BatchInsertCheckboxResult.Failed("No active workbook is open in Excel.");
            }

            Excel.Worksheet? activeWs = _app.ActiveSheet as Excel.Worksheet;
            if (activeWs == null)
            {
                return BatchInsertCheckboxResult.Failed("No active worksheet found.");
            }

            if (activeWs.ProtectContents)
            {
                return BatchInsertCheckboxResult.Failed($"Worksheet '{activeWs.Name}' is protected. Please unprotect it before inserting checkboxes.");
            }

            Excel.Range? sourceRange = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(options.SourceRangeAddress))
                {
                    sourceRange = activeWs.Range[options.SourceRangeAddress];
                }
                else
                {
                    sourceRange = _app.Selection as Excel.Range;
                }
            }
            catch (Exception ex)
            {
                return BatchInsertCheckboxResult.Failed($"Invalid source range address '{options.SourceRangeAddress}': {ex.Message}");
            }

            if (sourceRange == null)
            {
                return BatchInsertCheckboxResult.Failed("Please select a valid range of cells to insert checkboxes.");
            }

            Excel.Range? targetLinkStartCell = null;
            if (options.LinkMode == CheckboxLinkMode.OtherCells)
            {
                if (string.IsNullOrWhiteSpace(options.TargetLinkRangeAddress))
                {
                    return BatchInsertCheckboxResult.Failed("Please specify a target cell or range to link the checkboxes to.");
                }

                try
                {
                    targetLinkStartCell = _app.Range[options.TargetLinkRangeAddress];
                }
                catch (Exception ex)
                {
                    return BatchInsertCheckboxResult.Failed($"Invalid target link address '{options.TargetLinkRangeAddress}': {ex.Message}");
                }
            }

            int insertedCount = 0;
            var affectedAddresses = new List<string>();

            using (new ExcelStateScope(_app))
            {
                try
                {
                    Excel.CheckBoxes checkBoxes = (Excel.CheckBoxes)activeWs.CheckBoxes();
                    int sourceStartRow = sourceRange.Row;
                    int sourceStartCol = sourceRange.Column;
                    int rowCount = sourceRange.Rows.Count;
                    int colCount = sourceRange.Columns.Count;

                    int targetStartRow = targetLinkStartCell?.Row ?? sourceStartRow;
                    int targetStartCol = targetLinkStartCell?.Column ?? sourceStartCol;
                    Excel.Worksheet targetWs = (targetLinkStartCell != null && targetLinkStartCell.Worksheet != null)
                        ? (Excel.Worksheet)targetLinkStartCell.Worksheet
                        : activeWs;

                    for (int r = 1; r <= rowCount; r++)
                    {
                        for (int c = 1; c <= colCount; c++)
                        {
                            Excel.Range cell = (Excel.Range)sourceRange.Cells[r, c];
                            double cellLeft = Convert.ToDouble(cell.Left);
                            double cellTop = Convert.ToDouble(cell.Top);
                            double cellWidth = Convert.ToDouble(cell.Width);
                            double cellHeight = Convert.ToDouble(cell.Height);

                            string cellText = cell.Text?.ToString() ?? cell.Value2?.ToString() ?? string.Empty;
                            bool hasCaption = options.UseCellContentAsLabel && !string.IsNullOrWhiteSpace(cellText);

                            CheckboxEngine.CalculateCheckboxBounds(
                                cellLeft,
                                cellTop,
                                cellWidth,
                                cellHeight,
                                hasCaption,
                                out double left,
                                out double top,
                                out double width,
                                out double height);

                            Excel.CheckBox chk = (Excel.CheckBox)checkBoxes.Add(left, top, width, height);
                            chk.Placement = Excel.XlPlacement.xlMoveAndSize;

                            if (hasCaption)
                            {
                                chk.Caption = cellText.Trim();
                                if (options.ClearOriginalContentWhenLabeled && options.LinkMode != CheckboxLinkMode.OwnCell)
                                {
                                    cell.Value2 = null;
                                }
                            }
                            else
                            {
                                chk.Caption = string.Empty;
                            }

                            chk.Value = options.InitialChecked ? 1 : -4146; // xlOn : xlOff

                            if (options.LinkMode == CheckboxLinkMode.OwnCell)
                            {
                                string cellAddr = cell.get_Address(true, true, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                                chk.LinkedCell = cellAddr;
                                if (options.DisplayLogicalValue)
                                {
                                    cell.Value2 = options.InitialChecked;
                                }
                            }
                            else if (options.LinkMode == CheckboxLinkMode.OtherCells)
                            {
                                Excel.Range targetCell = (Excel.Range)targetWs.Cells[targetStartRow + (r - 1), targetStartCol + (c - 1)];
                                string targetAddr = targetCell.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true, Type.Missing);
                                chk.LinkedCell = targetAddr;
                                if (options.DisplayLogicalValue)
                                {
                                    targetCell.Value2 = options.InitialChecked;
                                }
                            }

                            string address = cell.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                            affectedAddresses.Add(address);
                            insertedCount++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    return BatchInsertCheckboxResult.Failed($"Failed while inserting checkboxes: {ex.Message}");
                }
            }

            return BatchInsertCheckboxResult.Successful(
                insertedCount,
                $"Successfully inserted {insertedCount} checkbox(es) into range.",
                affectedAddresses);
        }

        public BatchDeleteCheckboxResult DeleteCheckboxes(BatchDeleteCheckboxOptions options)
        {
            options ??= new BatchDeleteCheckboxOptions();

            if (!HasActiveWorkbook)
            {
                return BatchDeleteCheckboxResult.Failed("No active workbook is open in Excel.");
            }

            Excel.Worksheet? activeWs = _app.ActiveSheet as Excel.Worksheet;
            if (activeWs == null)
            {
                return BatchDeleteCheckboxResult.Failed("No active worksheet found.");
            }

            if (activeWs.ProtectContents)
            {
                return BatchDeleteCheckboxResult.Failed($"Worksheet '{activeWs.Name}' is protected. Please unprotect it before deleting checkboxes.");
            }

            int deletedCount = 0;
            using (new ExcelStateScope(_app))
            {
                try
                {
                    Excel.CheckBoxes checkBoxes = (Excel.CheckBoxes)activeWs.CheckBoxes();
                    if (checkBoxes == null || checkBoxes.Count == 0)
                    {
                        return BatchDeleteCheckboxResult.Successful(0, "No checkboxes were found on the worksheet.");
                    }

                    if (options.Scope == CheckboxDeleteScope.ActiveWorksheet)
                    {
                        deletedCount = checkBoxes.Count;
                        checkBoxes.Delete();
                    }
                    else
                    {
                        Excel.Range? targetRange = null;
                        if (!string.IsNullOrWhiteSpace(options.RangeAddress))
                        {
                            targetRange = activeWs.Range[options.RangeAddress];
                        }
                        else
                        {
                            targetRange = _app.Selection as Excel.Range;
                        }

                        if (targetRange == null)
                        {
                            return BatchDeleteCheckboxResult.Failed("No selection or range specified for deleting checkboxes.");
                        }

                        var toDelete = new List<Excel.CheckBox>();
                        foreach (Excel.CheckBox chk in checkBoxes)
                        {
                            try
                            {
                                Excel.Range? topLeft = chk.TopLeftCell;
                                if (topLeft != null)
                                {
                                    Excel.Range? intersect = _app.Intersect(topLeft, targetRange);
                                    if (intersect != null)
                                    {
                                        toDelete.Add(chk);
                                    }
                                }
                            }
                            catch { }
                        }

                        deletedCount = toDelete.Count;
                        foreach (var chk in toDelete)
                        {
                            try { chk.Delete(); } catch { }
                        }
                    }
                }
                catch (Exception ex)
                {
                    return BatchDeleteCheckboxResult.Failed($"Failed while deleting checkboxes: {ex.Message}");
                }
            }

            return BatchDeleteCheckboxResult.Successful(
                deletedCount,
                deletedCount == 0 ? "No checkboxes found in the designated scope." : $"Successfully deleted {deletedCount} checkbox(es).");
        }

        public int GetCheckboxCount(string? rangeAddress = null)
        {
            if (!HasActiveWorkbook) return 0;
            Excel.Worksheet? activeWs = _app.ActiveSheet as Excel.Worksheet;
            if (activeWs == null) return 0;

            try
            {
                Excel.CheckBoxes checkBoxes = (Excel.CheckBoxes)activeWs.CheckBoxes();
                if (checkBoxes == null || checkBoxes.Count == 0) return 0;

                if (string.IsNullOrWhiteSpace(rangeAddress))
                {
                    return checkBoxes.Count;
                }

                Excel.Range targetRange = activeWs.Range[rangeAddress];
                int count = 0;
                foreach (Excel.CheckBox chk in checkBoxes)
                {
                    try
                    {
                        Excel.Range? topLeft = chk.TopLeftCell;
                        if (topLeft != null && _app.Intersect(topLeft, targetRange) != null)
                        {
                            count++;
                        }
                    }
                    catch { }
                }
                return count;
            }
            catch
            {
                return 0;
            }
        }
    }
}
