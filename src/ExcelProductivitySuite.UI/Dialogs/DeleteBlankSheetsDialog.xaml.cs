using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class DeleteBlankSheetsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        public ObservableCollection<SelectableSheetItem> BlankItems { get; } = new();

        public DeleteBlankSheetsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            BlankSheetsListBox.ItemsSource = BlankItems;
            ScanBlankSheets();
        }

        private void ScanBlankSheets()
        {
            BlankItems.Clear();
            var blanks = _context.GetBlankWorksheetNames();

            foreach (var b in blanks)
            {
                BlankItems.Add(new SelectableSheetItem { Name = b, IsSelected = true });
            }

            if (BlankItems.Count == 0)
            {
                NoBlanksNotice.Visibility = Visibility.Visible;
                BlankSheetsListBox.Visibility = Visibility.Collapsed;
                DeleteButton.IsEnabled = false;
                StatusLabel.Text = "Zero blank worksheets detected in active workbook.";
                StatusLabel.Foreground = System.Windows.Media.Brushes.Gray;
            }
            else
            {
                NoBlanksNotice.Visibility = Visibility.Collapsed;
                BlankSheetsListBox.Visibility = Visibility.Visible;
                DeleteButton.IsEnabled = true;
                StatusLabel.Text = $"{BlankItems.Count} blank worksheet(s) found. Check items to remove.";
            }
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in BlankItems) item.IsSelected = true;
            BlankSheetsListBox.Items.Refresh();
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in BlankItems) item.IsSelected = false;
            BlankSheetsListBox.Items.Refresh();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var selected = BlankItems.Where(b => b.IsSelected).Select(b => b.Name).ToList();
            if (selected.Count == 0)
            {
                _context.ShowMessage("Please select at least one blank worksheet to delete.", "None Selected", isError: true);
                return;
            }

            bool confirmed = _context.Confirm(
                $"Are you sure you want to permanently delete {selected.Count} blank worksheet(s)?\n\nThis operation cannot be undone by Excel Undo.",
                "Confirm Sheet Deletion");

            if (!confirmed) return;

            int deleted = _context.DeleteWorksheets(selected);
            _context.ShowMessage($"Successfully deleted {deleted} blank worksheet(s).", "Sheets Deleted");
            DialogResult = true;
            Close();
        }
    }
}
