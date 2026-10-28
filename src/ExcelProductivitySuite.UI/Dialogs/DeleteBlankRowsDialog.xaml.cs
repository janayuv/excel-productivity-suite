using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class DeleteBlankRowsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;

        public DeleteBlankRowsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdateSummary();
        }

        private BlankRowDeleteOptions GetOptions()
        {
            var options = new BlankRowDeleteOptions();

            if (BlankColumnsRadio?.IsChecked == true)
            {
                options.DeleteBlankColumns = true;
            }
            else if (KeyColumnBlankRadio?.IsChecked == true)
            {
                options.CheckEntireRow = false;
                if (int.TryParse(KeyColBox?.Text?.Trim(), out int col))
                {
                    options.KeyColumnIndex = Math.Max(0, col - 1);
                }
            }
            else
            {
                options.CheckEntireRow = true;
            }

            return options;
        }

        private void UpdateSummary()
        {
            if (ScanSummaryText == null) return;
            if (_matrix == null)
            {
                ScanSummaryText.Text = "No range selected in Excel.";
                return;
            }

            try
            {
                var options = GetOptions();
                var result = BlankRowManager.DeleteBlankRows(_matrix, options, out int deleted);

                if (options.DeleteBlankColumns)
                {
                    int totalCols = _matrix.GetUpperBound(1) - _matrix.GetLowerBound(1) + 1;
                    ScanSummaryText.Text = $"Found {deleted} blank column(s) out of {totalCols} columns. {result.GetLength(1)} columns will remain.";
                }
                else
                {
                    int totalRows = _matrix.GetUpperBound(0) - _matrix.GetLowerBound(0) + 1;
                    ScanSummaryText.Text = $"Found {deleted} blank row(s) out of {totalRows} rows. {result.GetLength(0)} rows will remain.";
                }
            }
            catch (Exception ex)
            {
                ScanSummaryText.Text = "Scan error: " + ex.Message;
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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before deleting rows.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range.", "Multi-Area Selection", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Delete Blank Rows", isError: true);
                return;
            }

            var options = GetOptions();
            var compacted = BlankRowManager.DeleteBlankRows(_matrix, options, out int deleted);

            if (deleted == 0)
            {
                _context.ShowMessage("No blank rows/columns matched the selected criteria.", "Delete Blank Rows");
                return;
            }

            int startRow = _context.SelectionStartRow;
            int startCol = _context.SelectionStartColumn;

            _context.ClearSelection();
            if (compacted.GetLength(0) > 0 && compacted.GetLength(1) > 0)
            {
                _context.SetRangeValues(startRow, startCol, compacted);
            }

            DialogResult = true;
            Close();
            _context.ShowMessage($"Removed {deleted} blank {(options.DeleteBlankColumns ? "column(s)" : "row(s)")} successfully.", "Delete Blank Rows");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
