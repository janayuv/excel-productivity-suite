using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Worksheets;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class SplitDataToSheetsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;

        public SplitDataToSheetsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeInfoText.Text = $"Range: {_context.SelectionAddress ?? "No Selection"}";
            _matrix = _context.GetSelectedValues();

            PopulateColumns();
            UpdatePreview();
        }

        private void PopulateColumns()
        {
            if (_matrix == null || ColumnSelectorComboBox == null) return;

            ColumnSelectorComboBox.Items.Clear();
            int cStart = _matrix.GetLowerBound(1);
            int cEnd = _matrix.GetUpperBound(1);
            int rStart = _matrix.GetLowerBound(0);

            bool hasHeaders = HasHeadersCheckBox?.IsChecked == true;
            int startCol = _context.SelectionStartColumn;

            for (int c = cStart; c <= cEnd; c++)
            {
                int offset = c - cStart;
                string colLetter = GetColumnLetter(startCol + offset);
                string headerName = _matrix[rStart, c]?.ToString()?.Trim() ?? "";

                string label = hasHeaders && !string.IsNullOrEmpty(headerName)
                    ? $"{colLetter}: {headerName}"
                    : $"Column {colLetter}";

                ColumnSelectorComboBox.Items.Add(label);
            }

            if (ColumnSelectorComboBox.Items.Count > 0)
            {
                ColumnSelectorComboBox.SelectedIndex = 0;
            }
        }

        private static string GetColumnLetter(int columnNumber)
        {
            string colLetter = "";
            while (columnNumber > 0)
            {
                int mod = (columnNumber - 1) % 26;
                colLetter = (char)('A' + mod) + colLetter;
                columnNumber = (columnNumber - mod) / 26;
            }
            return string.IsNullOrEmpty(colLetter) ? "A" : colLetter;
        }

        private DataSplitOptions GetOptions()
        {
            int selectedCol = ColumnSelectorComboBox?.SelectedIndex ?? 0;
            if (selectedCol < 0) selectedCol = 0;

            return new DataSplitOptions
            {
                SplitColumnIndex = selectedCol,
                HasHeaders = HasHeadersCheckBox?.IsChecked == true,
                SheetNamePrefix = PrefixTextBox?.Text?.Trim() ?? ""
            };
        }

        private void UpdatePreview()
        {
            if (_matrix == null || PreviewSummaryText == null) return;

            try
            {
                var options = GetOptions();
                var result = DataSplitter.Split(_matrix, options);

                PreviewSummaryText.Text = $"Selected column contains {result.DistinctGroupsCount} distinct values.\nExecuting will create {result.DistinctGroupsCount} new worksheets across {result.TotalRowsProcessed} data rows.";
            }
            catch (Exception ex)
            {
                PreviewSummaryText.Text = $"Preview calculation: {ex.Message}";
            }
        }

        private void ColumnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void HeadersChanged(object sender, RoutedEventArgs e)
        {
            PopulateColumns();
            UpdatePreview();
        }

        private void PrefixChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void SplitButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsWorkbookProtected)
            {
                _context.ShowMessage("The active workbook structure is protected. Cannot create new worksheets.", "Workbook Protected", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected to split.", "Split Data", isError: true);
                return;
            }

            var options = GetOptions();
            var result = DataSplitter.Split(_matrix, options);

            if (result.DistinctGroupsCount == 0)
            {
                _context.ShowMessage("No distinct groups found to split.", "Split Data");
                return;
            }

            bool proceed = _context.Confirm($"Are you sure you want to split this dataset into {result.DistinctGroupsCount} new worksheets?\n\nOriginal worksheet data will remain untouched.", "Confirm Data Split");
            if (!proceed) return;

            int createdCount = 0;
            foreach (var kvp in result.SheetsToCreate)
            {
                bool ok = _context.CreateWorksheetWithData(kvp.Key, kvp.Value, activate: false);
                if (ok) createdCount++;
            }

            _context.ShowMessage($"Successfully created {createdCount} new worksheets from {result.TotalRowsProcessed} records.", "Split Data Completed");
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
