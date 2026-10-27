using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class MergeRangesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _range1Matrix;

        public MergeRangesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            Range1TextBox.Text = _context.SelectionAddress ?? "$A$1";
            _range1Matrix = _context.GetSelectedValues();
        }

        private RangeMergeOptions GetOptions()
        {
            return new RangeMergeOptions
            {
                Mode = ModeSideBySideRadio.IsChecked == true ? MergeMode.SideBySide : MergeMode.AppendRows,
                HasHeaders = HasHeadersCheckBox.IsChecked == true,
                AlignByHeaderNames = AlignHeadersCheckBox.IsChecked == true,
                IncludeHeadersOnce = IncludeHeadersOnceCheckBox.IsChecked == true
            };
        }

        private void MergeButton_Click(object sender, RoutedEventArgs e)
        {
            if (_range1Matrix == null)
            {
                _context.ShowMessage("No primary range cells selected.", "Merge Ranges", isError: true);
                return;
            }

            string targetAddress = Range2TextBox.Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(targetAddress))
            {
                _context.ShowMessage("Please specify the secondary range address to merge with.", "Merge Ranges", isError: true);
                return;
            }

            var range2Matrix = _context.GetRangeValuesByAddress(targetAddress);
            if (range2Matrix == null)
            {
                return; // Error already shown in adapter
            }

            var options = GetOptions();
            var result = RangeMerger.Merge(_range1Matrix, range2Matrix, options);

            if (result.MergedMatrix == null || result.TotalRows == 0)
            {
                _context.ShowMessage("Merge produced empty result.", "Merge Ranges", isError: true);
                return;
            }

            if (DestNewSheetRadio.IsChecked == true)
            {
                string sheetName = "Merged_Data";
                bool ok = _context.CreateWorksheetWithData(sheetName, result.MergedMatrix);
                if (ok)
                {
                    string warn = result.Warnings.Count > 0 ? $"\n\nNote: {string.Join("\n", result.Warnings)}" : "";
                    _context.ShowMessage($"Merged {result.TotalRows} rows and {result.TotalColumns} columns into new worksheet '{sheetName}'.{warn}", "Merge Completed");
                    DialogResult = true;
                    Close();
                }
            }
            else
            {
                if (_context.IsCurrentSheetProtected)
                {
                    _context.ShowMessage("The active worksheet is protected. Cannot write merged data.", "Sheet Protected", isError: true);
                    return;
                }

                int originRow = _context.SelectionStartRow;
                int originCol = _context.SelectionStartColumn + _context.SelectionColumnCount;

                if (_context.CheckRangeHasData(originRow, originCol, result.TotalRows, result.TotalColumns))
                {
                    bool proceed = _context.Confirm("The destination range already contains data. Merging will overwrite existing cells. Continue?", "Confirm Overwrite");
                    if (!proceed) return;
                }

                _context.SetRangeValues(originRow, originCol, result.MergedMatrix);
                _context.ShowMessage($"Merged data placed adjacent to selection ({result.TotalRows} rows, {result.TotalColumns} columns).", "Merge Completed");
                DialogResult = true;
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
