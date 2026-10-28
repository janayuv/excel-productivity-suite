using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formulas;
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

        public void ShowMessage(string message, string title, bool isError = false)
        {
            MessageBox.Show(message, title, MessageBoxButton.OK, isError ? MessageBoxImage.Error : MessageBoxImage.Information);
        }
    }
}
