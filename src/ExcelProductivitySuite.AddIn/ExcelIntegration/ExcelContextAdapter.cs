using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelProductivitySuite.AddIn.ExcelIntegration
{
    public class ExcelContextAdapter : IExcelContext
    {
        private readonly Excel.Application _app;

        public ExcelContextAdapter(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
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

            using (new ExcelStateScope(_app))
            {
                range.Value2 = values;
            }
        }

        public void SetRangeValues(int startRow, int startCol, object?[,] values)
        {
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (!(_app.ActiveSheet is Excel.Worksheet sheet)) return;

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            using (new ExcelStateScope(_app))
            {
                Excel.Range topLeft = sheet.Cells[startRow, startCol];
                Excel.Range bottomRight = sheet.Cells[startRow + rows - 1, startCol + cols - 1];
                Excel.Range targetRange = sheet.Range[topLeft, bottomRight];

                targetRange.Value2 = values;
            }
        }

        public int ConvertSelectionFormulasToValues()
        {
            if (!(_app.Selection is Excel.Range range)) return 0;

            using (new ExcelStateScope(_app))
            {
                range.Value2 = range.Value2;
            }

            return SelectedCellCount;
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

        public void ShowMessage(string message, string title, bool isError = false)
        {
            MessageBox.Show(message, title, MessageBoxButton.OK, isError ? MessageBoxImage.Error : MessageBoxImage.Information);
        }
    }
}
