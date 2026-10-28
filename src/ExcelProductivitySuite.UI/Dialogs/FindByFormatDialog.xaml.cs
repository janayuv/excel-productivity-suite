using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FindByFormatDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private CellFormatCriteria _currentCriteria = new CellFormatCriteria();
        private IReadOnlyList<FormatMatchItem> _matches = new List<FormatMatchItem>();

        public FindByFormatDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            PickActiveCellFormat();
        }

        private void PickActiveCellFormat()
        {
            _currentCriteria = _context.GetActiveCellFormatCriteria();
            MatchFillCheck.IsChecked = _currentCriteria.MatchFillColor;
            MatchFontColorCheck.IsChecked = _currentCriteria.MatchFontColor;
            MatchBoldCheck.IsChecked = _currentCriteria.MatchBold;
            MatchItalicCheck.IsChecked = _currentCriteria.MatchItalic;

            StatusText.Text = "Extracted formatting attributes from active cell. Ready to search.";
        }

        private void PickFromActiveCellBtn_Click(object sender, RoutedEventArgs e)
        {
            PickActiveCellFormat();
        }

        private void SearchBtn_Click(object sender, RoutedEventArgs e)
        {
            _currentCriteria.MatchFillColor = MatchFillCheck.IsChecked == true;
            _currentCriteria.MatchFontColor = MatchFontColorCheck.IsChecked == true;
            _currentCriteria.MatchBold = MatchBoldCheck.IsChecked == true;
            _currentCriteria.MatchItalic = MatchItalicCheck.IsChecked == true;

            bool selectionOnly = ScopeSelectionRadio.IsChecked == true;
            _matches = _context.FindCellsByFormat(_currentCriteria, selectionOnly);

            MatchesGrid.ItemsSource = _matches;
            SelectAllBtn.IsEnabled = _matches.Count > 0;

            StatusText.Text = _matches.Count > 0
                ? $"Found {_matches.Count} cell(s) matching selected formatting criteria."
                : "No matching cells found.";
        }

        private void SelectAllBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_matches == null || _matches.Count == 0) return;

            var addresses = _matches.Select(m => m.Address).ToList();
            _context.SelectCellsByAddresses(addresses);
            StatusText.Text = $"Selected {addresses.Count} matching cell(s) in Excel.";
        }

        private void MatchesGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (MatchesGrid.SelectedItem is FormatMatchItem item)
            {
                _context.SelectCellAddress(item.SheetName, item.Address);
                StatusText.Text = $"Navigated to {item.SheetName}!{item.Address} in Excel.";
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
