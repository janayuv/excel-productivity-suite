using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class RowDifferencesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;
        private List<(int Row, int Column, object? Value, object? BaselineValue)>? _differences;

        public RowDifferencesDialog(IExcelContext context, bool defaultToColumnDifferences = false)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            if (defaultToColumnDifferences && ColDiffRadio != null)
            {
                ColDiffRadio.IsChecked = true;
            }

            _matrix = _context.GetSelectedValues();
            UpdateSummary();
        }

        private DifferenceOptions GetOptions()
        {
            var options = new DifferenceOptions
            {
                MatchCase = MatchCaseCheck?.IsChecked == true,
                Direction = RowDiffRadio?.IsChecked == true
                    ? DifferenceDirection.RowDifferences
                    : DifferenceDirection.ColumnDifferences
            };

            if (int.TryParse(BaselineBox?.Text?.Trim(), out int idx) && idx > 0)
            {
                options.BaselineIndex = idx - 1; // 1-based to 0-based
            }

            return options;
        }

        private void UpdateSummary()
        {
            if (SummaryText == null) return;
            if (_matrix == null)
            {
                SummaryText.Text = "No range selected.";
                return;
            }

            try
            {
                var options = GetOptions();
                int startRow = _context.SelectionStartRow;
                int startCol = _context.SelectionStartColumn;

                _differences = RowDifferenceFinder.FindDifferences(_matrix, startRow, startCol, options);

                SummaryText.Text = $"Found {_differences.Count} difference cell(s) compared to baseline.";
            }
            catch (Exception ex)
            {
                SummaryText.Text = "Error: " + ex.Message;
            }
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdateSummary();
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (_differences == null || _differences.Count == 0)
            {
                _context.ShowMessage("No differences found matching criteria.", "Row Differences");
                return;
            }

            var coords = _differences.Select(c => (c.Row, c.Column)).ToList();
            _context.SelectCells(coords);

            DialogResult = true;
            Close();
            _context.ShowMessage($"Selected {coords.Count} differing cell(s) in Excel.", "Row Differences");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
