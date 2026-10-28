using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Worksheets;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class SortSheetTabsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _originalSheets = new();
        private List<string> _sortedSheets = new();

        public SortSheetTabsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            LoadSheets();
        }

        private void LoadSheets()
        {
            _originalSheets.Clear();
            _originalSheets.AddRange(_context.GetWorksheetNames());

            CurrentOrderListBox.ItemsSource = _originalSheets;
            UpdatePreview();
        }

        private SheetSortMode GetSortMode()
        {
            if (AlphaAscRadio?.IsChecked == true) return SheetSortMode.AlphabeticalAscending;
            if (AlphaDescRadio?.IsChecked == true) return SheetSortMode.AlphabeticalDescending;
            if (ReverseRadio?.IsChecked == true) return SheetSortMode.ReverseCurrent;
            return SheetSortMode.NaturalNumeric;
        }

        private void UpdatePreview()
        {
            var mode = GetSortMode();
            _sortedSheets = SheetTabSorter.Sort(_originalSheets, mode);
            PreviewOrderListBox.ItemsSource = null;
            PreviewOrderListBox.ItemsSource = _sortedSheets;

            if (SummaryLabel != null)
            {
                SummaryLabel.Text = $"{_originalSheets.Count} worksheet tabs ready to reorder.";
            }
        }

        private void SortMode_Changed(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ApplySort_Click(object sender, RoutedEventArgs e)
        {
            if (_sortedSheets.Count == 0)
            {
                Close();
                return;
            }

            _context.ReorderWorksheets(_sortedSheets);
            _context.ShowMessage($"Successfully reordered {_sortedSheets.Count} worksheet tabs.", "Tabs Sorted");
            DialogResult = true;
            Close();
        }
    }
}
