using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class WorkbookFileItem
    {
        public string FullPath { get; set; } = "";
        public string FileName => Path.GetFileName(FullPath);
    }

    public partial class CombineWorkbooksDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        public ObservableCollection<WorkbookFileItem> Files { get; } = new();

        public CombineWorkbooksDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            FilesListBox.ItemsSource = Files;
            UpdateStatus();
        }

        private void UpdateStatus()
        {
            StatusLabel.Text = $"{Files.Count} workbook file(s) queued for import.";
        }

        private void AddFiles_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new System.Windows.Forms.OpenFileDialog())
            {
                dlg.Title = "Select Excel Workbooks to Combine";
                dlg.Filter = "Excel Files (*.xlsx;*.xlsm;*.xls;*.csv)|*.xlsx;*.xlsm;*.xls;*.csv|All Files (*.*)|*.*";
                dlg.Multiselect = true;

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (var f in dlg.FileNames)
                    {
                        if (!Files.Any(item => string.Equals(item.FullPath, f, StringComparison.OrdinalIgnoreCase)))
                        {
                            Files.Add(new WorkbookFileItem { FullPath = f });
                        }
                    }
                    UpdateStatus();
                }
            }
        }

        private void RemoveFile_Click(object sender, RoutedEventArgs e)
        {
            if (FilesListBox.SelectedItem is WorkbookFileItem item)
            {
                Files.Remove(item);
                UpdateStatus();
            }
        }

        private void ClearFiles_Click(object sender, RoutedEventArgs e)
        {
            Files.Clear();
            UpdateStatus();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Combine_Click(object sender, RoutedEventArgs e)
        {
            if (Files.Count == 0)
            {
                _context.ShowMessage("Please add at least one Excel workbook to combine.", "No Files", isError: true);
                return;
            }

            bool firstOnly = FirstSheetRadio.IsChecked == true;
            bool prefix = PrefixNamesCheck.IsChecked == true;

            int totalImported = 0;
            foreach (var f in Files)
            {
                int count = _context.ImportSheetsFromWorkbook(f.FullPath, firstOnly, prefix);
                totalImported += count;
            }

            _context.ShowMessage($"Successfully imported {totalImported} worksheet(s) from {Files.Count} workbook(s).", "Workbooks Combined");
            DialogResult = true;
            Close();
        }
    }
}
