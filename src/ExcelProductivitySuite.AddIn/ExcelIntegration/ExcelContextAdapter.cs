using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
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

        public void ShowMessage(string message, string title, bool isError = false)
        {
            MessageBox.Show(message, title, MessageBoxButton.OK, isError ? MessageBoxImage.Error : MessageBoxImage.Information);
        }
    }
}
