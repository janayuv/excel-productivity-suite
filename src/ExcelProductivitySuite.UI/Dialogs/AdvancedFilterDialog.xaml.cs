using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class FilteredRowDisplayItem
    {
        public int RowNumber { get; set; }
        public string PreviewSummary { get; set; } = "";
    }

    public partial class AdvancedFilterDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;
        private AdvancedFilterResult? _lastResult;

        public AdvancedFilterDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeInfoText.Text = $"Range: {_context.SelectionAddress ?? "No Selection"}";
            _matrix = _context.GetSelectedValues();

            PopulateOperators();
            PopulateColumns();
            UpdateFilterPreview();
        }

        private void PopulateOperators()
        {
            OperatorComboBox.Items.Clear();
            foreach (FilterOperator op in Enum.GetValues(typeof(FilterOperator)))
            {
                OperatorComboBox.Items.Add(op.ToString());
            }
            OperatorComboBox.SelectedItem = FilterOperator.Equals.ToString();
        }

        private void PopulateColumns()
        {
            if (_matrix == null || ColumnComboBox == null) return;

            ColumnComboBox.Items.Clear();
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

                ColumnComboBox.Items.Add(label);
            }

            if (ColumnComboBox.Items.Count > 0)
            {
                ColumnComboBox.SelectedIndex = 0;
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

        private AdvancedFilterOptions GetOptions()
        {
            int colIndex = ColumnComboBox?.SelectedIndex ?? 0;
            if (colIndex < 0) colIndex = 0;

            FilterOperator op = FilterOperator.Equals;
            if (OperatorComboBox?.SelectedItem is string opStr)
            {
                Enum.TryParse(opStr, out op);
            }

            var condition = new FilterCondition
            {
                ColumnIndex = colIndex,
                Operator = op,
                ComparisonValue = ComparisonValueTextBox?.Text ?? ""
            };

            return new AdvancedFilterOptions
            {
                Conjunction = ConjunctionOrRadio?.IsChecked == true ? FilterConjunction.Or : FilterConjunction.And,
                HasHeaders = HasHeadersCheckBox?.IsChecked == true,
                Conditions = new List<FilterCondition> { condition }
            };
        }

        private void UpdateFilterPreview()
        {
            if (_matrix == null || MatchingRecordsDataGrid == null) return;

            var options = GetOptions();
            _lastResult = AdvancedFilter.Execute(_matrix, options);

            var items = _lastResult.MatchingRows.Select((row, idx) => new FilteredRowDisplayItem
            {
                RowNumber = _lastResult.MatchingRowIndices[idx] + 1,
                PreviewSummary = string.Join(" | ", row.Select(v => v?.ToString() ?? ""))
            }).ToList();

            MatchingRecordsDataGrid.ItemsSource = items;
        }

        private void FilterCriteriaChanged(object sender, RoutedEventArgs e)
        {
            UpdateFilterPreview();
        }

        private void HeadersChanged(object sender, RoutedEventArgs e)
        {
            PopulateColumns();
            UpdateFilterPreview();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null || _lastResult.MatchCount == 0)
            {
                _context.ShowMessage("No rows matched the filter criteria.", "Advanced Filter");
                return;
            }

            if (ActionExportNewSheetRadio.IsChecked == true)
            {
                if (_context.IsWorkbookProtected)
                {
                    _context.ShowMessage("The active workbook structure is protected. Cannot create new worksheet.", "Workbook Protected", isError: true);
                    return;
                }

                var matrix = _lastResult.ExportToMatrix();
                string sheetName = "Filtered_Data";
                bool ok = _context.CreateWorksheetWithData(sheetName, matrix);

                if (ok)
                {
                    _context.ShowMessage($"Exported {_lastResult.MatchCount} matching rows to worksheet '{sheetName}'.", "Filter Export Complete");
                    DialogResult = true;
                    Close();
                }
            }
            else // Select first match in Excel
            {
                int firstRowOffset = _lastResult.MatchingRowIndices[0];
                int targetSheetRow = _context.SelectionStartRow + firstRowOffset;
                int startCol = _context.SelectionStartColumn;
                int colCount = _context.SelectionColumnCount;

                _context.SelectRange(targetSheetRow, startCol, 1, colCount);
                _context.ShowMessage($"Selected first matching row in Excel ({_lastResult.MatchCount} total rows matched).", "Filter Navigation");
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
