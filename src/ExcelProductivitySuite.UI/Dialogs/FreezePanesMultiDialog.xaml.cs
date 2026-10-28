using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FreezePanesMultiDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly ObservableCollection<SheetSelectionItem> _targetSheets = new();

        public FreezePanesMultiDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SheetsList.ItemsSource = _targetSheets;
            PopulateSheets();
        }

        private void PopulateSheets()
        {
            _targetSheets.Clear();
            var allSheets = _context.GetWorksheetNames();
            foreach (var s in allSheets)
            {
                _targetSheets.Add(new SheetSelectionItem { Name = s, IsSelected = true });
            }
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _targetSheets)
            {
                item.IsSelected = true;
            }
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _targetSheets)
            {
                item.IsSelected = false;
            }
        }

        private FreezePanesSyncMode SelectedMode
        {
            get
            {
                if (FreezeFirstColRadio?.IsChecked == true) return FreezePanesSyncMode.FreezeFirstColumn;
                if (FreezeActiveCellRadio?.IsChecked == true) return FreezePanesSyncMode.FreezeAtActiveCell;
                if (UnfreezeAllRadio?.IsChecked == true) return FreezePanesSyncMode.UnfreezeAll;
                return FreezePanesSyncMode.FreezeTopRow;
            }
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            var targets = _targetSheets.Where(s => s.IsSelected).Select(s => s.Name).ToList();
            if (targets.Count == 0)
            {
                StatusLabel.Text = "Please select at least one worksheet.";
                return;
            }

            int splitRow = _context.SelectionStartRow;
            int splitCol = _context.SelectionStartColumn;

            try
            {
                int synchronized = _context.SynchronizeFreezePanes(targets, SelectedMode, splitRow, splitCol);
                StatusLabel.Text = $"Synchronized freeze panes across {synchronized} worksheet(s).";
                _context.ShowMessage($"Updated freeze panes on {synchronized} worksheet(s).", "Freeze Panes Multi");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error: {ex.Message}";
                _context.ShowMessage($"Failed to synchronize freeze panes: {ex.Message}", "Freeze Panes Error", isError: true);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
