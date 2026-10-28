using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class LocateFormulaErrorsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private List<FormulaErrorInfo> _allErrors = new();

        public LocateFormulaErrorsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            PopulateFilterCombo();
            ScanErrors();
        }

        private void PopulateFilterCombo()
        {
            ErrorTypeCombo.Items.Clear();
            ErrorTypeCombo.Items.Add("All Errors");
            ErrorTypeCombo.Items.Add("#DIV/0!");
            ErrorTypeCombo.Items.Add("#N/A");
            ErrorTypeCombo.Items.Add("#VALUE!");
            ErrorTypeCombo.Items.Add("#REF!");
            ErrorTypeCombo.Items.Add("#NAME?");
            ErrorTypeCombo.Items.Add("#NUM!");
            ErrorTypeCombo.Items.Add("#NULL!");
            ErrorTypeCombo.SelectedIndex = 0;
        }

        private void ScanErrors()
        {
            bool selOnly = SelectionScopeRadio?.IsChecked == true;
            _allErrors = _context.GetFormulaErrors(selOnly).ToList();
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (ErrorsGrid == null) return;

            string selectedFilter = ErrorTypeCombo?.SelectedItem?.ToString() ?? "All Errors";
            IEnumerable<FormulaErrorInfo> filtered = _allErrors;

            if (!string.Equals(selectedFilter, "All Errors", StringComparison.OrdinalIgnoreCase))
            {
                filtered = filtered.Where(e => string.Equals(e.ErrorText, selectedFilter, StringComparison.OrdinalIgnoreCase));
            }

            var list = filtered.ToList();
            ErrorsGrid.ItemsSource = list;
            StatusLabel.Text = $"Found {list.Count} error cell(s)";
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            if (IsLoaded)
            {
                if (sender == SheetScopeRadio || sender == SelectionScopeRadio)
                {
                    ScanErrors();
                }
                else
                {
                    ApplyFilter();
                }
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            ScanErrors();
        }

        private void ErrorsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Optional auto-focus
        }

        private void GoToCell_Click(object sender, RoutedEventArgs e)
        {
            if (ErrorsGrid.SelectedItem is FormulaErrorInfo err)
            {
                _context.SelectRange(err.Row, err.Column, 1, 1);
            }
            else
            {
                StatusLabel.Text = "Please select an error row from the table first.";
            }
        }

        private void SelectAllErrors_Click(object sender, RoutedEventArgs e)
        {
            var list = ErrorsGrid.ItemsSource as List<FormulaErrorInfo>;
            if (list == null || list.Count == 0)
            {
                StatusLabel.Text = "No errors to select.";
                return;
            }

            var cells = list.Select(err => (err.Row, err.Column)).ToList();
            _context.SelectCells(cells);
            StatusLabel.Text = $"Selected {cells.Count} error cell(s) in Excel.";
        }

        private void HighlightErrors_Click(object sender, RoutedEventArgs e)
        {
            var list = ErrorsGrid.ItemsSource as List<FormulaErrorInfo>;
            if (list == null || list.Count == 0)
            {
                StatusLabel.Text = "No errors to highlight.";
                return;
            }

            var cells = list.Select(err => (err.Row, err.Column)).ToList();
            // Light coral red: RGB 0xF8D7DA = 255, 215, 218
            int colorRgb = 0xF8D7DA;
            _context.HighlightCells(cells, colorRgb);
            StatusLabel.Text = $"Highlighted {cells.Count} error cell(s).";
            _context.ShowMessage($"Applied highlight fill to {cells.Count} error cell(s).", "Highlight Errors");
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
