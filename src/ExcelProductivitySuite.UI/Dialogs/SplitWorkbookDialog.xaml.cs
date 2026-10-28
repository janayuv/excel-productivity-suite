using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Worksheets;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class SplitWorkbookDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        public ObservableCollection<SelectableSheetItem> SheetItems { get; } = new();

        public SplitWorkbookDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SheetsListBox.ItemsSource = SheetItems;
            FolderBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
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

            StatusLabel.Text = $"{SheetItems.Count} sheets loaded.";
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

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new System.Windows.Forms.FolderBrowserDialog())
            {
                dlg.Description = "Select Destination Folder for Exported Sheets";
                dlg.SelectedPath = Directory.Exists(FolderBox.Text) ? FolderBox.Text : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    FolderBox.Text = dlg.SelectedPath;
                }
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Split_Click(object sender, RoutedEventArgs e)
        {
            var selected = SheetItems.Where(s => s.IsSelected).Select(s => s.Name).ToList();
            if (selected.Count == 0)
            {
                _context.ShowMessage("Please select at least one worksheet to export.", "No Sheets Selected", isError: true);
                return;
            }

            string targetDir = FolderBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(targetDir) || !Directory.Exists(targetDir))
            {
                try { Directory.CreateDirectory(targetDir); }
                catch (Exception ex)
                {
                    _context.ShowMessage($"Invalid destination directory: {ex.Message}", "Invalid Path", isError: true);
                    return;
                }
            }

            var format = SplitFileFormat.ExcelWorkbook;
            if (PdfRadio.IsChecked == true) format = SplitFileFormat.Pdf;
            else if (CsvRadio.IsChecked == true) format = SplitFileFormat.Csv;

            string pattern = string.IsNullOrWhiteSpace(PatternBox.Text) ? "[SheetName]" : PatternBox.Text.Trim();

            int exported = _context.SplitWorksheetsToFiles(selected, targetDir, format, pattern);
            if (exported > 0)
            {
                _context.ShowMessage($"Successfully exported {exported} worksheet(s) to:\n{targetDir}", "Workbook Split Complete");
                DialogResult = true;
                Close();
            }
        }
    }
}
