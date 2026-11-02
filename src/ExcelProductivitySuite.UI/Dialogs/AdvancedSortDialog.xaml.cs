using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class AdvancedSortDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;

        public AdvancedSortDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            PopulateColumns();
        }

        private void PopulateColumns()
        {
            if (ColumnCombo == null) return;
            ColumnCombo.Items.Clear();
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
                string headerName = hasHeaders ? _matrix[rStart, cStart + c]?.ToString() ?? "" : "";
                if (string.IsNullOrWhiteSpace(headerName))
                {
                    headerName = $"Column {c + 1}";
                }
                ColumnCombo.Items.Add(headerName);
            }

            if (ColumnCombo.Items.Count > 0)
            {
                ColumnCombo.SelectedIndex = 0;
            }

            int rows = _matrix.GetUpperBound(0) - rStart + 1;
            if (StatusLabel != null)
            {
                StatusLabel.Text = $"{rows} rows x {cols} columns in selection.";
            }
        }

        private void HasHeaders_Changed(object sender, RoutedEventArgs e)
        {
            PopulateColumns();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Sort_Click(object sender, RoutedEventArgs e)
        {
            if (_matrix == null)
            {
                Close();
                return;
            }

            var criterion = AdvancedSortCriterion.TextLengthAscending;
            if (LenDescRadio.IsChecked == true) criterion = AdvancedSortCriterion.TextLengthDescending;
            else if (MonthRadio.IsChecked == true) criterion = AdvancedSortCriterion.MonthName;
            else if (WeekdayRadio.IsChecked == true) criterion = AdvancedSortCriterion.WeekdayName;
            else if (StdAscRadio.IsChecked == true) criterion = AdvancedSortCriterion.StandardAscending;
            else if (StdDescRadio.IsChecked == true) criterion = AdvancedSortCriterion.StandardDescending;

            var options = new AdvancedSortOptions
            {
                SortColumnIndex = Math.Max(0, ColumnCombo.SelectedIndex),
                HasHeaderRow = HasHeadersCheck.IsChecked == true,
                Criterion = criterion
            };

            var sorted = AdvancedSorter.Sort(_matrix, options);
            _context.SetSelectedValues(sorted);
            _context.ShowMessage("Selection successfully sorted by specified rule.", "Advanced Sort Complete");
            DialogResult = true;
            Close();
        }
    }
}
