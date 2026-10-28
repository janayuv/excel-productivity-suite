using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FormulaAuditDependentsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public class DependentItemViewModel
        {
            public string SheetName { get; }
            public string CellAddress { get; }
            public string ValueText { get; }
            public string FormulaText { get; }
            public bool IsDirect { get; }
            public string FormulaDisplay => string.IsNullOrWhiteSpace(FormulaText) ? "(Unknown Formula)" : FormulaText;

            public DependentItemViewModel(CellReferenceInfo info)
            {
                SheetName = info.SheetName;
                CellAddress = info.CellAddress;
                ValueText = string.IsNullOrWhiteSpace(info.ValueText) ? "(Empty)" : info.ValueText;
                FormulaText = info.FormulaText;
                IsDirect = info.IsDirect;
            }
        }

        public ObservableCollection<DependentItemViewModel> DependentsList { get; } = new ObservableCollection<DependentItemViewModel>();

        public FormulaAuditDependentsDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            DependentsGrid.ItemsSource = DependentsList;
            LoadDependents();
        }

        private void LoadDependents()
        {
            DependentsList.Clear();
            NavigateBtn.IsEnabled = false;

            var cellDetails = _context.GetActiveCellDetails();
            if (cellDetails == null)
            {
                ActiveCellText.Text = "None";
                ActiveValueText.Text = "No active cell selected.";
                StatusText.Text = "Please select a cell in Excel.";
                return;
            }

            ActiveCellText.Text = $"{cellDetails.SheetName}!{cellDetails.Address}";
            ActiveValueText.Text = string.IsNullOrWhiteSpace(cellDetails.ValueText) ? "(Empty)" : cellDetails.ValueText;

            bool directOnly = DirectOnlyRadio.IsChecked == true;
            var dependents = _context.GetActiveCellDependents(directOnly);

            foreach (var d in dependents)
            {
                DependentsList.Add(new DependentItemViewModel(d));
            }

            StatusText.Text = DependentsList.Count > 0
                ? $"Found {DependentsList.Count} dependent cell(s) that calculate from active cell."
                : "No dependent cells detected (no formulas rely on this cell).";
        }

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (IsLoaded) LoadDependents();
        }

        private void DependentsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            NavigateBtn.IsEnabled = DependentsGrid.SelectedItem is DependentItemViewModel;
        }

        private void DependentsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            NavigateToSelectedDependent();
        }

        private void NavigateBtn_Click(object sender, RoutedEventArgs e)
        {
            NavigateToSelectedDependent();
        }

        private void NavigateToSelectedDependent()
        {
            if (DependentsGrid.SelectedItem is DependentItemViewModel item)
            {
                bool success = _context.SelectCellAddress(item.SheetName, item.CellAddress);
                if (success)
                {
                    StatusText.Text = $"Navigated to {item.SheetName}!{item.CellAddress} in Excel.";
                }
                else
                {
                    StatusText.Text = $"Could not navigate to {item.SheetName}!{item.CellAddress}.";
                }
            }
        }

        private void RefreshBtn_Click(object sender, RoutedEventArgs e)
        {
            LoadDependents();
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
