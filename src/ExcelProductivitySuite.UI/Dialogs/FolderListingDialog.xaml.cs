using System;
using System.IO;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Files;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FolderListingDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public FolderListingDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            FolderPathText.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            using var fbd = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select Folder to Inventory",
                SelectedPath = FolderPathText.Text,
                ShowNewFolderButton = false
            };

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                FolderPathText.Text = fbd.SelectedPath;
            }
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            string folder = FolderPathText.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                StatusLabel.Text = "Please select a valid folder path.";
                return;
            }

            try
            {
                var options = new FolderScanOptions
                {
                    IncludeSubfolders = IncludeSubfoldersCheck.IsChecked == true,
                    SearchPattern = string.IsNullOrWhiteSpace(FilterPatternText.Text) ? "*.*" : FilterPatternText.Text.Trim(),
                    GenerateHyperlinks = HyperlinksCheck.IsChecked == true
                };

                var scanResult = FolderInventoryScanner.ScanFolder(folder, options);
                if (scanResult.FileCount == 0)
                {
                    StatusLabel.Text = "No files found matching the criteria.";
                    _context.ShowMessage("No files found matching the specified pattern.", "Folder Listing");
                    return;
                }

                if (NewSheetRadio.IsChecked == true)
                {
                    string sheetName = NewSheetNameText.Text?.Trim() ?? "File_Listing";
                    _context.CreateWorksheetWithData(sheetName, scanResult.TableData);
                    _context.ShowMessage($"Generated directory listing for {scanResult.FileCount} file(s) ({FolderInventoryScanner.FormatBytes(scanResult.TotalBytes)}) in new sheet '{sheetName}'.", "Folder Listing");
                }
                else
                {
                    int startRow = _context.SelectionStartRow;
                    int startCol = _context.SelectionStartColumn;
                    _context.SetRangeValues(startRow, startCol, scanResult.TableData);
                    _context.ShowMessage($"Inserted {scanResult.FileCount} file(s) into current sheet.", "Folder Listing");
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error: {ex.Message}";
                _context.ShowMessage($"Failed to generate listing: {ex.Message}", "Folder Listing Error", isError: true);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
