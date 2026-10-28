using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formulas;
using ExcelProductivitySuite.Engine.Navigation;
using ExcelProductivitySuite.Engine.Ranges;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelProductivitySuite.AddIn.ExcelIntegration
{
    public class ExcelContextAdapter : IExcelContext
    {
        private readonly Func<Excel.Application> _appAccessor;
        private Excel.Application _app => _appAccessor() ?? throw new InvalidOperationException("Excel Application instance is not available.");

        public ExcelContextAdapter(Excel.Application app)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));
            _appAccessor = () => app;
        }

        public ExcelContextAdapter(Func<Excel.Application> appAccessor)
        {
            _appAccessor = appAccessor ?? throw new ArgumentNullException(nameof(appAccessor));
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
                var matrix = GetSelectedValues();
                if (matrix == null) return samples;

                int rStart = matrix.GetLowerBound(0);
                int rEnd = matrix.GetUpperBound(0);
                int cStart = matrix.GetLowerBound(1);

                for (int r = rStart; r <= rEnd && samples.Count < maxSamples; r++)
                {
                    var val = matrix[r, cStart];
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
                targetRange.Formula = formulas;
            }
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

        public bool ExportToPdf(string destinationPath, IEnumerable<string>? sheetNames = null, bool openAfterExport = false)
        {
            if (string.IsNullOrWhiteSpace(destinationPath)) throw new ArgumentNullException(nameof(destinationPath));

            try
            {
                if (!(_app.ActiveWorkbook is Excel.Workbook wb)) return false;

                using (new ExcelStateScope(_app))
                {
                    var sheetList = sheetNames != null ? new List<string>(sheetNames) : new List<string>();

                    if (sheetList.Count == 1)
                    {
                        string sName = sheetList[0];
                        Excel.Worksheet? targetWs = null;
                        foreach (Excel.Worksheet ws in wb.Worksheets)
                        {
                            if (string.Equals(ws.Name, sName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetWs = ws;
                                break;
                            }
                        }

                        if (targetWs != null)
                        {
                            targetWs.ExportAsFixedFormat(
                                Excel.XlFixedFormatType.xlTypePDF,
                                destinationPath,
                                Excel.XlFixedFormatQuality.xlQualityStandard,
                                IncludeDocProperties: true,
                                IgnorePrintAreas: false,
                                OpenAfterPublish: openAfterExport);
                            return true;
                        }
                    }

                    // Export entire workbook
                    wb.ExportAsFixedFormat(
                        Excel.XlFixedFormatType.xlTypePDF,
                        destinationPath,
                        Excel.XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        OpenAfterPublish: openAfterExport);

                    return true;
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error exporting PDF: {ex.Message}", "PDF Export Error", isError: true);
                return false;
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
    }
}
