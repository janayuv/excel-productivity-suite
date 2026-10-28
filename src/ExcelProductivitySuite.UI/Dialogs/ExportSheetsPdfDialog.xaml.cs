using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;
using Microsoft.Win32;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ExportSheetsPdfDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly ObservableCollection<SheetSelectionItem> _sheetItems = new();

        public ExportSheetsPdfDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SheetsList.ItemsSource = _sheetItems;
            PopulateSheets();

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string baseName = _context.ActiveWorksheetName ?? "Workbook";
            DestinationPathText.Text = Path.Combine(desktop, $"{baseName}.pdf");
        }

        private void PopulateSheets()
        {
            _sheetItems.Clear();
            var allSheets = _context.GetWorksheetNames();
            foreach (var s in allSheets)
            {
                _sheetItems.Add(new SheetSelectionItem { Name = s, IsSelected = true });
            }
        }

        private void Scope_Changed(object sender, RoutedEventArgs e)
        {
            if (SheetsBox != null)
            {
                SheetsBox.Visibility = SelectedSheetsRadio?.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _sheetItems) item.IsSelected = true;
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _sheetItems) item.IsSelected = false;
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var sfd = new SaveFileDialog
            {
                Title = "Save PDF As",
                Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*",
                FileName = Path.GetFileName(DestinationPathText.Text)
            };

            if (sfd.ShowDialog() == true)
            {
                DestinationPathText.Text = sfd.FileName;
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            string dest = DestinationPathText.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(dest))
            {
                StatusLabel.Text = "Please enter a valid destination PDF path.";
                return;
            }

            try
            {
                bool openAfter = OpenAfterExportCheck.IsChecked == true;
                bool success;

                if (ActiveSheetRadio.IsChecked == true)
                {
                    string sheet = _context.ActiveWorksheetName ?? "";
                    success = _context.ExportToPdf(dest, new[] { sheet }, openAfter);
                }
                else if (SelectedSheetsRadio.IsChecked == true)
                {
                    var targets = _sheetItems.Where(s => s.IsSelected).Select(s => s.Name).ToList();
                    if (targets.Count == 0)
                    {
                        StatusLabel.Text = "Please select at least one worksheet.";
                        return;
                    }
                    success = _context.ExportToPdf(dest, targets, openAfter);
                }
                else
                {
                    // Entire workbook
                    success = _context.ExportToPdf(dest, null, openAfter);
                }

                if (success)
                {
                    StatusLabel.Text = "PDF published successfully!";
                    _context.ShowMessage($"Exported PDF to '{dest}'.", "Export Complete");
                    DialogResult = true;
                    Close();
                }
                else
                {
                    StatusLabel.Text = "Failed to export PDF.";
                }
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Export error: {ex.Message}";
                _context.ShowMessage($"Export failed: {ex.Message}", "PDF Error", isError: true);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
