using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class CellDisplayItem
    {
        public string Description { get; set; } = "";
        public int Row { get; set; }
        public int Column { get; set; }
        public string Value { get; set; } = "";
    }

    public partial class SelectSpecialDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;
        private List<CellCoordinate> _matchedCells = new List<CellCoordinate>();

        public SelectSpecialDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeInfoText.Text = $"Range: {_context.SelectionAddress ?? "No Selection"}";
            _matrix = _context.GetSelectedValues();

            UpdateMatchingCells();
        }

        private SelectionCriteria GetCriteria()
        {
            if (RadioBlank.IsChecked == true) return SelectionCriteria.BlankCells;
            if (RadioNonBlank.IsChecked == true) return SelectionCriteria.NonBlankCells;
            if (RadioErrors.IsChecked == true) return SelectionCriteria.ErrorCells;
            if (RadioFormulas.IsChecked == true) return SelectionCriteria.FormulaCells;
            if (RadioNumeric.IsChecked == true) return SelectionCriteria.NumericCells;
            if (RadioText.IsChecked == true) return SelectionCriteria.TextCells;
            if (RadioMax.IsChecked == true) return SelectionCriteria.MaxNumericValue;
            if (RadioMin.IsChecked == true) return SelectionCriteria.MinNumericValue;
            return SelectionCriteria.BlankCells;
        }

        private void CriteriaChanged(object sender, RoutedEventArgs e)
        {
            UpdateMatchingCells();
        }

        private void UpdateMatchingCells()
        {
            if (_matrix == null || MatchingCellsGrid == null) return;

            var criteria = GetCriteria();
            int startRow = _context.SelectionStartRow;
            int startCol = _context.SelectionStartColumn;

            _matchedCells = RangeSelector.FindMatchingCells(_matrix, startRow, startCol, criteria);

            var items = _matchedCells.Select(c => new CellDisplayItem
            {
                Description = $"${GetColumnLetter(c.Column)}${c.Row}",
                Row = c.Row,
                Column = c.Column,
                Value = c.Value?.ToString() ?? "(Blank)"
            }).ToList();

            MatchingCellsGrid.ItemsSource = items;
            CountSummaryText.Text = $"Found {_matchedCells.Count} matching cells.";

            if (SelectInExcelButton != null)
            {
                SelectInExcelButton.IsEnabled = _matchedCells.Count > 0;
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

        private void SelectInExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (_matchedCells.Count > 0)
            {
                var first = _matchedCells[0];
                _context.SelectRange(first.Row, first.Column, 1, 1);
                _context.ShowMessage($"Selected first matching cell: ${GetColumnLetter(first.Column)}${first.Row} ({_matchedCells.Count} total found).", "Select Special");
                Close();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
