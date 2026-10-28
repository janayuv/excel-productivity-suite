using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class InsertBlankRowsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;

        public InsertBlankRowsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdateSummary();
        }

        private BlankRowInsertOptions GetOptions()
        {
            var options = new BlankRowInsertOptions
            {
                HasHeaderRow = HasHeaderCheck?.IsChecked == true
            };

            if (OnValueChangeRadio?.IsChecked == true)
            {
                options.Mode = BlankRowInsertMode.OnValueChange;
                if (int.TryParse(KeyColBox?.Text?.Trim(), out int col))
                {
                    options.KeyColumnIndex = Math.Max(0, col - 1); // 1-based to 0-based
                }
                if (int.TryParse(ChangeBlankCountBox?.Text?.Trim(), out int bCount))
                {
                    options.BlankRowCount = Math.Max(1, bCount);
                }
            }
            else
            {
                options.Mode = BlankRowInsertMode.EveryNRows;
                if (int.TryParse(IntervalBox?.Text?.Trim(), out int interval))
                {
                    options.Interval = Math.Max(1, interval);
                }
                if (int.TryParse(BlankCountBox?.Text?.Trim(), out int bCount))
                {
                    options.BlankRowCount = Math.Max(1, bCount);
                }
            }

            return options;
        }

        private void UpdateSummary()
        {
            if (SummaryInfoText == null) return;
            if (_matrix == null)
            {
                SummaryInfoText.Text = "No range selected in Excel.";
                return;
            }

            try
            {
                var options = GetOptions();
                var result = BlankRowManager.InsertBlankRows(_matrix, options);

                int origRows = _matrix.GetUpperBound(0) - _matrix.GetLowerBound(0) + 1;
                int newRows = result.GetLength(0);
                int added = newRows - origRows;

                SummaryInfoText.Text = $"Source: {origRows} rows  ➔  Target: {newRows} rows (+{added} blank rows inserted).";
            }
            catch (Exception ex)
            {
                SummaryInfoText.Text = "Rule error: " + ex.Message;
            }
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdateSummary();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before inserting rows.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range.", "Multi-Area Selection", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Insert Blank Rows", isError: true);
                return;
            }

            var options = GetOptions();
            var expanded = BlankRowManager.InsertBlankRows(_matrix, options);

            int startRow = _context.SelectionStartRow;
            int startCol = _context.SelectionStartColumn;

            _context.SetRangeValues(startRow, startCol, expanded);
            DialogResult = true;
            Close();

            int added = expanded.GetLength(0) - (_matrix.GetUpperBound(0) - _matrix.GetLowerBound(0) + 1);
            _context.ShowMessage($"Inserted {added} blank row(s) into range.", "Insert Blank Rows");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
