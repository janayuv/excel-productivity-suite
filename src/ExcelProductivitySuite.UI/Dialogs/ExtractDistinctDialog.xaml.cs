using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class SelectableColumnItem
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public bool IsSelected { get; set; } = true;
    }

    public partial class ExtractDistinctDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;
        public ObservableCollection<SelectableColumnItem> ColumnItems { get; } = new();

        public ExtractDistinctDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            ColumnsListBox.ItemsSource = ColumnItems;
            PopulateColumns();
        }

        private void PopulateColumns()
        {
            if (ColumnsListBox == null) return;
            ColumnItems.Clear();
            if (_matrix == null)
            {
                if (StatusLabel != null) StatusLabel.Text = "No range selected.";
                return;
            }

            int rStart = _matrix.GetLowerBound(0);
            int cStart = _matrix.GetLowerBound(1);
            int cEnd = _matrix.GetUpperBound(1);
            int cols = cEnd - cStart + 1;
            bool hasHeaders = HasHeadersCheck?.IsChecked == true;

            for (int c = 0; c < cols; c++)
            {
                string name = hasHeaders ? _matrix[rStart, cStart + c]?.ToString() ?? "" : "";
                if (string.IsNullOrWhiteSpace(name)) name = $"Column {c + 1}";
                ColumnItems.Add(new SelectableColumnItem { Index = c, Name = name, IsSelected = true });
            }

            int rows = _matrix.GetUpperBound(0) - rStart + 1;
            if (StatusLabel != null)
            {
                StatusLabel.Text = $"{rows} rows x {cols} columns ready to process.";
            }
        }

        private void HasHeaders_Changed(object sender, RoutedEventArgs e)
        {
            PopulateColumns();
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in ColumnItems) item.IsSelected = true;
            ColumnsListBox.Items.Refresh();
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in ColumnItems) item.IsSelected = false;
            ColumnsListBox.Items.Refresh();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Extract_Click(object sender, RoutedEventArgs e)
        {
            if (_matrix == null)
            {
                Close();
                return;
            }

            var selectedCols = ColumnItems.Where(c => c.IsSelected).Select(c => c.Index).ToArray();
            if (selectedCols.Length == 0)
            {
                _context.ShowMessage("Please select at least one column to evaluate for uniqueness.", "No Columns", isError: true);
                return;
            }

            var options = new DistinctExtractOptions
            {
                HasHeaders = HasHeadersCheck.IsChecked == true,
                KeyColumns = selectedCols,
                MatchCase = MatchCaseCheck.IsChecked == true
            };

            var distinct = DistinctExtractor.ExtractDistinct(_matrix, options, out int uniqueCount, out int duplicateCount);

            string sheetName = string.IsNullOrWhiteSpace(DestSheetBox.Text) ? "Distinct_Records" : DestSheetBox.Text.Trim();
            bool success = _context.CreateWorksheetWithData(sheetName, distinct, activate: true);

            if (success)
            {
                _context.ShowMessage($"Extracted {uniqueCount} distinct record(s) to worksheet '{sheetName}' ({duplicateCount} duplicate rows omitted).", "Extraction Complete");
                DialogResult = true;
                Close();
            }
        }
    }
}
