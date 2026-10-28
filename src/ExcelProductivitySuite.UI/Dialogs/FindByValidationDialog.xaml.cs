using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FindByValidationDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private IReadOnlyList<ValidationRuleInfo> _allRules = new List<ValidationRuleInfo>();
        private List<ValidationRuleInfo> _filteredRules = new List<ValidationRuleInfo>();

        public FindByValidationDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadRules();
        }

        private void LoadRules()
        {
            _allRules = _context.GetWorksheetValidationRules(selectionOnly: false);
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (FilterListsRadio.IsChecked == true)
            {
                _filteredRules = _allRules.Where(r => r.ValidationType.Equals("List", StringComparison.OrdinalIgnoreCase)).ToList();
            }
            else if (FilterInvalidRadio.IsChecked == true)
            {
                _filteredRules = _allRules.Where(r => r.IsValid == false).ToList();
            }
            else
            {
                _filteredRules = _allRules.ToList();
            }

            RulesGrid.ItemsSource = _filteredRules;
            bool hasItems = _filteredRules.Count > 0;
            SelectAllBtn.IsEnabled = hasItems;
            RemoveValBtn.IsEnabled = hasItems;

            StatusText.Text = hasItems
                ? $"Displaying {_filteredRules.Count} of {_allRules.Count} data validation cell(s)."
                : "No data validation cells match the selected filter.";
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            if (IsLoaded) ApplyFilter();
        }

        private void RefreshBtn_Click(object sender, RoutedEventArgs e)
        {
            LoadRules();
        }

        private void SelectAllBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_filteredRules.Count == 0) return;

            var addresses = _filteredRules.Select(r => r.CellAddress).ToList();
            _context.SelectCellsByAddresses(addresses);
            StatusText.Text = $"Selected {addresses.Count} validation cell(s) in Excel.";
        }

        private void RemoveValBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_filteredRules.Count == 0) return;

            bool confirm = _context.Confirm(
                $"Remove validation rules from {_filteredRules.Count} cell(s)?",
                "Confirm Validation Removal");

            if (!confirm) return;

            var addresses = _filteredRules.Select(r => r.CellAddress).ToList();
            _context.RemoveValidationFromRange(addresses);
            _context.ShowMessage($"Successfully removed validation rules from {addresses.Count} cell(s).", "Rules Removed");
            LoadRules();
        }

        private void RulesGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (RulesGrid.SelectedItem is ValidationRuleInfo item)
            {
                _context.SelectCellAddress(item.SheetName, item.CellAddress);
                StatusText.Text = $"Navigated to {item.SheetName}!{item.CellAddress} in Excel.";
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
