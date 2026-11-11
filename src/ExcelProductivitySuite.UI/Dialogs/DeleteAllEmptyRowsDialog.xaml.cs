using System;
using System.Windows;
using System.Windows.Media;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    /// <summary>
    /// Interaction logic for DeleteAllEmptyRowsDialog.xaml.
    /// Implements ASAP Utilities Tool #232 (Columns &amp; Rows &gt; Delete All Empty Rows).
    /// </summary>
    public partial class DeleteAllEmptyRowsDialog : BaseDialogWindow
    {
        private readonly IExcelContext? _context;
        private EmptyRowsScanResult? _lastScanResult;

        public DeleteAllEmptyRowsDialog()
        {
            InitializeComponent();
        }

        public DeleteAllEmptyRowsDialog(IExcelContext context) : this()
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeDialog();
        }

        private void InitializeDialog()
        {
            if (_context == null) return;

            string activeSheet = _context.ActiveWorksheetName ?? "Active Sheet";
            if (ScopeActiveSheetRadio != null)
            {
                ScopeActiveSheetRadio.Content = $"Active Worksheet ('{activeSheet}') - Entire Dataset";
            }

            if (ScopeSelectionRadio != null)
            {
                string selAddr = _context.SelectionAddress ?? "Selection";
                ScopeSelectionRadio.Content = $"Current Selection ({selAddr})";
            }

            UpdateScanSummary();
        }

        private DeleteEmptyRowsOptions GetOptions()
        {
            var options = new DeleteEmptyRowsOptions
            {
                TreatWhitespaceAsEmpty = ChkTreatWhitespace?.IsChecked == true,
                TreatEmptyFormulaAsEmpty = ChkTreatEmptyFormulas?.IsChecked == true,
                ConfirmBeforeDelete = false
            };

            if (ScopeSelectionRadio?.IsChecked == true)
            {
                options.Scope = DeleteEmptyRowsScope.CurrentSelection;
            }
            else if (ScopeAllSheetsRadio?.IsChecked == true)
            {
                options.Scope = DeleteEmptyRowsScope.AllWorksheetsInWorkbook;
            }
            else
            {
                options.Scope = DeleteEmptyRowsScope.ActiveWorksheet;
            }

            return options;
        }

        private void UpdateScanSummary()
        {
            if (_context == null || ScanSummaryHeader == null) return;

            try
            {
                var options = GetOptions();

                if (options.Scope == DeleteEmptyRowsScope.AllWorksheetsInWorkbook)
                {
                    var sheetNames = _context.GetWorksheetNames();
                    int totalSheets = sheetNames.Count;
                    _lastScanResult = _context.ScanEmptyRows(options);

                    ScanSummaryHeader.Text = $"All Worksheets ({totalSheets} sheets)";
                    ScanDetailsText.Text = $"Will scan and remove completely empty rows across all {totalSheets} worksheets in the active workbook.";

                    int activeEmpty = _lastScanResult.EmptyRowCount;
                    EmptyCountText.Text = activeEmpty > 0 ? $"{activeEmpty}+" : "0+";
                    RemainingCountText.Text = "Bulk";
                    RangeBreakdownText.Text = $"Active sheet '{_lastScanResult.WorksheetName}': {activeEmpty} empty row(s) found.";

                    DeleteButton.IsEnabled = true;
                    StatusBadge.Foreground = new SolidColorBrush(Color.FromRgb(46, 125, 50));
                    return;
                }

                _lastScanResult = _context.ScanEmptyRows(options);

                if (_lastScanResult == null || _lastScanResult.TotalRowsScanned == 0)
                {
                    ScanSummaryHeader.Text = "No data rows detected";
                    ScanDetailsText.Text = "The target range or worksheet has no occupied rows or is completely blank.";
                    EmptyCountText.Text = "0";
                    RemainingCountText.Text = "0";
                    RangeBreakdownText.Text = "No rows to delete.";
                    DeleteButton.IsEnabled = false;
                    StatusBadge.Foreground = new SolidColorBrush(Color.FromRgb(158, 158, 158));
                    return;
                }

                int emptyCount = _lastScanResult.EmptyRowCount;
                int totalScanned = _lastScanResult.TotalRowsScanned;
                int remaining = _lastScanResult.RemainingRowCount;

                EmptyCountText.Text = emptyCount.ToString("N0");
                RemainingCountText.Text = remaining.ToString("N0");

                if (emptyCount == 0)
                {
                    ScanSummaryHeader.Text = "Worksheet dataset is already compact";
                    ScanDetailsText.Text = $"Scanned {totalScanned} row(s) in '{_lastScanResult.WorksheetName}'. No completely empty rows were found.";
                    RangeBreakdownText.Text = "✓ No empty rows found. No deletion required.";
                    DeleteButton.IsEnabled = false;
                    StatusBadge.Foreground = new SolidColorBrush(Color.FromRgb(46, 125, 50));
                }
                else
                {
                    ScanSummaryHeader.Text = $"Found {emptyCount} completely empty row(s)";
                    ScanDetailsText.Text = $"Scanned {totalScanned} row(s) in '{_lastScanResult.WorksheetName}'. Deleting these empty gaps will leave {remaining} clean data rows.";
                    RangeBreakdownText.Text = $"Row range(s) to remove: {_lastScanResult.FormattedRowRanges}";
                    DeleteButton.IsEnabled = true;
                    StatusBadge.Foreground = new SolidColorBrush(Color.FromRgb(211, 47, 47));
                }
            }
            catch (Exception ex)
            {
                ScanSummaryHeader.Text = "Scan error";
                ScanDetailsText.Text = ex.Message;
                DeleteButton.IsEnabled = false;
                StatusBadge.Foreground = new SolidColorBrush(Color.FromRgb(211, 47, 47));
            }
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdateScanSummary();
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context == null) return;

            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before deleting rows.", "Sheet Protected", isError: true);
                return;
            }

            var options = GetOptions();

            // Confirmation guard for safety
            int emptyCount = _lastScanResult?.EmptyRowCount ?? 0;
            string scopeDesc = options.Scope switch
            {
                DeleteEmptyRowsScope.CurrentSelection => "the current selection",
                DeleteEmptyRowsScope.AllWorksheetsInWorkbook => "all worksheets in the active workbook",
                _ => $"active worksheet '{_context.ActiveWorksheetName}'"
            };

            string confirmPrompt = options.Scope == DeleteEmptyRowsScope.AllWorksheetsInWorkbook
                ? $"Are you sure you want to delete all completely empty rows across ALL worksheets in the active workbook?\n\nThis will shift non-empty rows up."
                : $"Are you sure you want to delete {emptyCount} completely empty row(s) in {scopeDesc}?\n\nThis will shift non-empty rows up.";

            if (!_context.Confirm(confirmPrompt, "Confirm Delete All Empty Rows"))
            {
                return;
            }

            var result = _context.DeleteAllEmptyRows(options);

            if (result.Success)
            {
                DialogResult = true;
                Close();
                _context.ShowMessage(result.Message, "Delete All Empty Rows - ASAP Utilities #232");
            }
            else
            {
                _context.ShowMessage(result.ErrorMessage ?? "Row deletion failed.", "Delete All Empty Rows", isError: true);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
