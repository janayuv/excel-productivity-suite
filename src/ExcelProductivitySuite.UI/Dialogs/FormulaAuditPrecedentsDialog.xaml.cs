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
    public partial class FormulaAuditPrecedentsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public class PrecedentItemViewModel
        {
            public string SheetName { get; }
            public string CellAddress { get; }
            public string ValueText { get; }
            public string FormulaText { get; }
            public bool IsDirect { get; }
            public string FormulaDisplay => string.IsNullOrWhiteSpace(FormulaText) ? "(Constant Value)" : FormulaText;

            public PrecedentItemViewModel(CellReferenceInfo info)
            {
                SheetName = info.SheetName;
                CellAddress = info.CellAddress;
                ValueText = string.IsNullOrWhiteSpace(info.ValueText) ? "(Empty)" : info.ValueText;
                FormulaText = info.FormulaText;
                IsDirect = info.IsDirect;
            }
        }

        public ObservableCollection<PrecedentItemViewModel> PrecedentsList { get; } = new ObservableCollection<PrecedentItemViewModel>();

        public FormulaAuditPrecedentsDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            PrecedentsGrid.ItemsSource = PrecedentsList;
            LoadPrecedents();
        }

        private void LoadPrecedents()
        {
            PrecedentsList.Clear();
            NavigateBtn.IsEnabled = false;

            var cellDetails = _context.GetActiveCellDetails();
            if (cellDetails == null)
            {
                ActiveCellText.Text = "None";
                ActiveFormulaText.Text = "No active cell selected.";
                StatusText.Text = "Please select a cell in Excel.";
                return;
            }

            ActiveCellText.Text = $"{cellDetails.SheetName}!{cellDetails.Address}";
            ActiveFormulaText.Text = string.IsNullOrWhiteSpace(cellDetails.FormulaText) ? "(No Formula)" : cellDetails.FormulaText;

            if (string.IsNullOrWhiteSpace(cellDetails.FormulaText) || !cellDetails.FormulaText.StartsWith("="))
            {
                StatusText.Text = "Active cell contains a static constant, not a formula.";
                return;
            }

            bool directOnly = DirectOnlyRadio.IsChecked == true;
            var precedents = _context.GetActiveCellPrecedents(directOnly);

            foreach (var p in precedents)
            {
                PrecedentsList.Add(new PrecedentItemViewModel(p));
            }

            StatusText.Text = PrecedentsList.Count > 0
                ? $"Found {PrecedentsList.Count} precedent cell(s) feeding into active formula."
                : "No precedent cells detected (formula may contain only scalar literals).";
        }

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (IsLoaded) LoadPrecedents();
        }

        private void PrecedentsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            NavigateBtn.IsEnabled = PrecedentsGrid.SelectedItem is PrecedentItemViewModel;
        }

        private void PrecedentsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            NavigateToSelectedPrecedent();
        }

        private void NavigateBtn_Click(object sender, RoutedEventArgs e)
        {
            NavigateToSelectedPrecedent();
        }

        private void NavigateToSelectedPrecedent()
        {
            if (PrecedentsGrid.SelectedItem is PrecedentItemViewModel item)
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
            LoadPrecedents();
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
