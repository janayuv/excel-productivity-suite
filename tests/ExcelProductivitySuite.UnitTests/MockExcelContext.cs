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

        public object?[,]? CurrentValues { get; set; }
        public List<string> MessagesShown { get; } = new();
        public List<string> Sheets { get; } = new() { "Sheet1", "Sheet2", "Sheet3" };
        public Dictionary<string, bool> SheetVisibility { get; } = new()
        {
            { "Sheet1", true },
            { "Sheet2", false },
            { "Sheet3", true }
        };

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

        public void ShowMessage(string message, string title, bool isError = false)
        {
            MessagesShown.Add($"[{title}] {message}");
        }
    }
}
