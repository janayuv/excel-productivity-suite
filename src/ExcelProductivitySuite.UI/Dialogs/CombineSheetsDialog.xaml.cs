using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Worksheets;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class SelectableSheetItem
    {
        public string Name { get; set; } = "";
        public bool IsSelected { get; set; } = true;
    }

    public partial class CombineSheetsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        public ObservableCollection<SelectableSheetItem> SheetItems { get; } = new();

        public CombineSheetsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SheetsListBox.ItemsSource = SheetItems;
            LoadSheets();
        }

        private void LoadSheets()
        {
            SheetItems.Clear();
            var names = _context.GetWorksheetNames();
            foreach (var n in names)
            {
                SheetItems.Add(new SelectableSheetItem { Name = n, IsSelected = true });
            }

            StatusLabel.Text = $"{SheetItems.Count} worksheets loaded.";
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in SheetItems) item.IsSelected = true;
            SheetsListBox.Items.Refresh();
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in SheetItems) item.IsSelected = false;
            SheetsListBox.Items.Refresh();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Combine_Click(object sender, RoutedEventArgs e)
        {
            var selected = SheetItems.Where(s => s.IsSelected).Select(s => s.Name).ToList();
            if (selected.Count == 0)
            {
                _context.ShowMessage("Please select at least one worksheet to combine.", "No Sheets Selected", isError: true);
                return;
            }

            int.TryParse(BlankRowsBox.Text, out int blankRows);
            string masterName = string.IsNullOrWhiteSpace(MasterSheetNameBox.Text) ? "Combined_Data" : MasterSheetNameBox.Text.Trim();

            var options = new SheetCombineOptions
            {
                HasHeaders = HasHeadersCheck.IsChecked == true,
                KeepFirstSheetHeaderOnly = KeepFirstHeaderCheck.IsChecked == true,
                AddSourceSheetColumn = AddSourceColCheck.IsChecked == true,
                BlankRowsBetweenSheets = Math.Max(0, blankRows)
            };

            var dataList = new List<SheetDataInfo>();
            foreach (var sheetName in selected)
            {
                var data = _context.GetActiveWorksheetRangeValues(sheetName);
                if (data != null)
                {
                    dataList.Add(new SheetDataInfo(sheetName, data));
                }
            }

            if (dataList.Count == 0)
            {
                _context.ShowMessage("None of the selected worksheets contain data.", "No Data", isError: true);
                return;
            }

            var combined = SheetCombiner.Combine(dataList, options, out int rowsCombined);
            bool success = _context.CreateWorksheetWithData(masterName, combined, activate: true);

            if (success)
            {
                _context.ShowMessage($"Successfully combined {selected.Count} worksheets into '{masterName}' ({rowsCombined} rows total).", "Sheets Combined");
                DialogResult = true;
                Close();
            }
        }
    }
}
