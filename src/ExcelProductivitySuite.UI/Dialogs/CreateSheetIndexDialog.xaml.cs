using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class SheetIndexDisplayItem
    {
        public string Name { get; set; } = "";
        public bool IsVisible { get; set; } = true;
        public string StatusText => IsVisible ? "(Visible)" : "(Hidden)";
    }

    public partial class CreateSheetIndexDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<SheetVisibilityInfo> _allSheets = new();
        public ObservableCollection<SheetIndexDisplayItem> DisplayItems { get; } = new();

        public CreateSheetIndexDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SheetsListBox.ItemsSource = DisplayItems;
            LoadSheets();
        }

        private void LoadSheets()
        {
            _allSheets.Clear();
            _allSheets.AddRange(_context.GetWorksheetVisibilities());
            RefreshList();
        }

        private void RefreshList()
        {
            DisplayItems.Clear();
            bool includeHidden = IncludeHiddenCheck?.IsChecked == true;

            foreach (var s in _allSheets)
            {
                if (!includeHidden && !s.IsVisible) continue;
                DisplayItems.Add(new SheetIndexDisplayItem
                {
                    Name = s.SheetName,
                    IsVisible = s.IsVisible
                });
            }

            if (StatusLabel != null)
            {
                StatusLabel.Text = $"{DisplayItems.Count} sheets will be linked in the Table of Contents.";
            }
        }

        private void Option_Changed(object sender, RoutedEventArgs e)
        {
            RefreshList();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CreateIndex_Click(object sender, RoutedEventArgs e)
        {
            string indexName = string.IsNullOrWhiteSpace(IndexNameBox.Text) ? "Table of Contents" : IndexNameBox.Text.Trim();
            var targetSheetNames = DisplayItems.Select(d => d.Name).ToList();

            if (targetSheetNames.Count == 0)
            {
                _context.ShowMessage("No worksheets available to index.", "No Sheets", isError: true);
                return;
            }

            bool includeBackLinks = AddBackLinksCheck.IsChecked == true;
            bool success = _context.CreateIndexWorksheet(indexName, targetSheetNames, includeBackLinks);

            if (success)
            {
                _context.ShowMessage($"Table of Contents created as '{indexName}' with {targetSheetNames.Count} clickable sheet links.", "Index Created");
                DialogResult = true;
                Close();
            }
        }
    }
}
