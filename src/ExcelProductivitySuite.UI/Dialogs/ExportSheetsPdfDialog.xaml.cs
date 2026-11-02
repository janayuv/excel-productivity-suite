using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Files;
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
            InitializeControls();
        }

        private void InitializeControls()
        {
            // Populate ComboBox options
            OrientationCombo.Items.Add("Default (Worksheet Setting)");
            OrientationCombo.Items.Add("Portrait");
            OrientationCombo.Items.Add("Landscape");
            OrientationCombo.SelectedIndex = 0;

            PageFitCombo.Items.Add("Default (No Scaling / As Is)");
            PageFitCombo.Items.Add("Fit Sheet on One Page");
            PageFitCombo.Items.Add("Fit All Columns on One Page Wide");
            PageFitCombo.Items.Add("Fit All Rows on One Page Tall");
            PageFitCombo.SelectedIndex = 0;

            PaperSizeCombo.Items.Add("Default (Worksheet Setting)");
            PaperSizeCombo.Items.Add("A4 (210 x 297 mm)");
            PaperSizeCombo.Items.Add("Letter (8.5 x 11 in)");
            PaperSizeCombo.Items.Add("Legal (8.5 x 14 in)");
            PaperSizeCombo.Items.Add("A3 (297 x 420 mm)");
            PaperSizeCombo.Items.Add("Tabloid (11 x 17 in)");
            PaperSizeCombo.Items.Add("Executive");
            PaperSizeCombo.SelectedIndex = 0;

            QualityCombo.Items.Add("Standard (High Quality Print & Screen)");
            QualityCombo.Items.Add("Minimum Size (Web & Email)");
            QualityCombo.SelectedIndex = 0;

            PatternCombo.Items.Add("[WorkbookName] - [SheetName]");
            PatternCombo.Items.Add("[SheetName]");
            PatternCombo.Items.Add("[Index] - [SheetName]");
            PatternCombo.Items.Add("[WorkbookName]_[SheetName]_[Date]");
            PatternCombo.Text = "[WorkbookName] - [SheetName]";

            // Populate sheets
            var allSheets = _context.GetWorksheetNames();
            _sheetItems.Clear();
            foreach (var s in allSheets)
            {
                _sheetItems.Add(new SheetSelectionItem { Name = s, IsSelected = true });
            }

            // Scope Labels
            string activeSheet = _context.ActiveWorksheetName ?? "Sheet1";
            ActiveSheetRadio.Content = $"Active Worksheet Only ({activeSheet})";

            string? selAddress = _context.SelectionAddress;
            if (!string.IsNullOrWhiteSpace(selAddress))
            {
                SelectionRadio.Content = $"Selection Only ({selAddress})";
                SelectionRadio.IsEnabled = true;
            }
            else
            {
                SelectionRadio.Content = "Selection Only (No Selection)";
                SelectionRadio.IsEnabled = false;
            }

            EntireWorkbookRadio.Content = $"Entire Active Workbook ({allSheets.Count} sheets)";

            // Default Paths
            string defaultFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (!string.IsNullOrWhiteSpace(_context.ActiveWorkbookPath))
            {
                string? wbDir = Path.GetDirectoryName(_context.ActiveWorkbookPath);
                if (!string.IsNullOrWhiteSpace(wbDir) && Directory.Exists(wbDir))
                {
                    defaultFolder = wbDir;
                }
            }

            OutputFolderText.Text = defaultFolder;

            string baseName = !string.IsNullOrWhiteSpace(_context.ActiveWorkbookName)
                ? Path.GetFileNameWithoutExtension(_context.ActiveWorkbookName)
                : activeSheet;

            DestinationPathText.Text = Path.Combine(defaultFolder, $"{baseName}.pdf");
        }

        private void Scope_Changed(object sender, RoutedEventArgs e)
        {
            if (SheetsBox != null)
            {
                SheetsBox.Visibility = SelectedSheetsRadio?.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
            }

            if (SelectionRadio?.IsChecked == true)
            {
                if (SeparatePdfRadio != null) SeparatePdfRadio.IsEnabled = false;
                if (SinglePdfRadio != null) SinglePdfRadio.IsChecked = true;
            }
            else
            {
                if (SeparatePdfRadio != null) SeparatePdfRadio.IsEnabled = true;
            }

            // Update default filename based on selected scope
            if (DestinationPathText != null && !string.IsNullOrWhiteSpace(DestinationPathText.Text))
            {
                string dir = Path.GetDirectoryName(DestinationPathText.Text) ?? "";
                string cleanWb = Path.GetFileNameWithoutExtension(_context.ActiveWorkbookName ?? "Workbook");
                string activeSheet = _context.ActiveWorksheetName ?? "Sheet1";

                string newName = cleanWb;
                if (ActiveSheetRadio?.IsChecked == true) newName = activeSheet;
                else if (SelectionRadio?.IsChecked == true) newName = $"{activeSheet}_Selection";
                else if (SelectedSheetsRadio?.IsChecked == true) newName = $"{cleanWb}_SelectedSheets";

                DestinationPathText.Text = Path.Combine(dir, $"{PdfExportEngine.SanitizeFileName(newName)}.pdf");
            }
        }

        private void OutputMode_Changed(object sender, RoutedEventArgs e)
        {
            if (SingleFilePanel == null || SeparateFilesPanel == null) return;

            bool isSingle = SinglePdfRadio?.IsChecked == true;
            SingleFilePanel.Visibility = isSingle ? Visibility.Visible : Visibility.Collapsed;
            SeparateFilesPanel.Visibility = isSingle ? Visibility.Collapsed : Visibility.Visible;
        }

        private void PageRange_Changed(object sender, RoutedEventArgs e)
        {
            if (PageFromText == null || PageToText == null) return;

            bool isRange = PageRangeRadio?.IsChecked == true;
            PageFromText.IsEnabled = isRange;
            PageToText.IsEnabled = isRange;
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _sheetItems) item.IsSelected = true;
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _sheetItems) item.IsSelected = false;
        }

        private void InvertSelection_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _sheetItems) item.IsSelected = !item.IsSelected;
        }

        private void BrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var sfd = new SaveFileDialog
            {
                Title = "Save Destination PDF",
                Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*",
                FileName = Path.GetFileName(DestinationPathText.Text)
            };

            if (sfd.ShowDialog() == true)
            {
                DestinationPathText.Text = sfd.FileName;
            }
        }

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            using var fbd = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select destination folder for exported PDF files",
                SelectedPath = OutputFolderText.Text
            };

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                OutputFolderText.Text = fbd.SelectedPath;
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var settings = new PdfExportSettings();

                // 1. Scope
                if (ActiveSheetRadio.IsChecked == true)
                {
                    settings.Scope = PdfExportScope.ActiveSheet;
                }
                else if (SelectionRadio.IsChecked == true)
                {
                    settings.Scope = PdfExportScope.Selection;
                }
                else if (SelectedSheetsRadio.IsChecked == true)
                {
                    settings.Scope = PdfExportScope.SelectedSheets;
                    settings.SelectedSheets = _sheetItems.Where(s => s.IsSelected).Select(s => s.Name).ToList();

                    if (settings.SelectedSheets.Count == 0)
                    {
                        StatusLabel.Text = "Please select at least one worksheet.";
                        _context.ShowMessage("Please select at least one worksheet to export.", "PDF Export", isError: true);
                        return;
                    }
                }
                else
                {
                    settings.Scope = PdfExportScope.EntireWorkbook;
                }

                // 2. Output Mode & Paths
                if (SinglePdfRadio.IsChecked == true)
                {
                    settings.OutputMode = PdfOutputMode.SingleCombinedFile;
                    settings.DestinationPath = DestinationPathText.Text?.Trim() ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(settings.DestinationPath))
                    {
                        StatusLabel.Text = "Please enter a valid destination PDF file path.";
                        _context.ShowMessage("Please enter a valid destination PDF file path.", "PDF Export", isError: true);
                        return;
                    }
                }
                else
                {
                    settings.OutputMode = PdfOutputMode.SeparateFilesPerSheet;
                    settings.OutputDirectory = OutputFolderText.Text?.Trim() ?? string.Empty;
                    settings.FileNamePattern = string.IsNullOrWhiteSpace(PatternCombo.Text)
                        ? "[WorkbookName] - [SheetName]"
                        : PatternCombo.Text.Trim();

                    if (string.IsNullOrWhiteSpace(settings.OutputDirectory))
                    {
                        StatusLabel.Text = "Please enter an output folder path.";
                        _context.ShowMessage("Please enter an output folder path.", "PDF Export", isError: true);
                        return;
                    }
                }

                // 3. Page Layout Overrides
                settings.Orientation = OrientationCombo.SelectedIndex switch
                {
                    1 => PdfPageOrientation.Portrait,
                    2 => PdfPageOrientation.Landscape,
                    _ => PdfPageOrientation.Default
                };

                settings.PageFit = PageFitCombo.SelectedIndex switch
                {
                    1 => PdfPageFit.FitSheetOnOnePage,
                    2 => PdfPageFit.FitAllColumnsOnOnePage,
                    3 => PdfPageFit.FitAllRowsOnOnePage,
                    _ => PdfPageFit.Default
                };

                settings.PaperSize = PaperSizeCombo.SelectedIndex switch
                {
                    1 => PdfPaperSize.A4,
                    2 => PdfPaperSize.Letter,
                    3 => PdfPaperSize.Legal,
                    4 => PdfPaperSize.A3,
                    5 => PdfPaperSize.Tabloid,
                    6 => PdfPaperSize.Executive,
                    _ => PdfPaperSize.Default
                };

                // 4. Page Range
                if (PageRangeRadio.IsChecked == true)
                {
                    settings.ExportPageRange = true;
                    if (!int.TryParse(PageFromText.Text?.Trim(), out int fromPage) || fromPage < 1)
                    {
                        StatusLabel.Text = "From page number must be 1 or greater.";
                        _context.ShowMessage("From page number must be 1 or greater.", "PDF Export", isError: true);
                        return;
                    }

                    if (!int.TryParse(PageToText.Text?.Trim(), out int toPage) || toPage < fromPage)
                    {
                        StatusLabel.Text = "To page number cannot be less than From page number.";
                        _context.ShowMessage("To page number cannot be less than From page number.", "PDF Export", isError: true);
                        return;
                    }

                    settings.PageFrom = fromPage;
                    settings.PageTo = toPage;
                }
                else
                {
                    settings.ExportPageRange = false;
                }

                // 5. Quality & Options
                settings.Quality = QualityCombo.SelectedIndex == 1 ? PdfQuality.MinimumSize : PdfQuality.Standard;
                settings.IgnorePrintAreas = IgnorePrintAreasCheck.IsChecked == true;
                settings.IncludeDocumentProperties = IncludeDocPropsCheck.IsChecked == true;
                settings.OpenAfterExport = OpenAfterExportCheck.IsChecked == true;
                settings.OpenOutputFolder = OpenFolderCheck.IsChecked == true;

                // Execute
                StatusLabel.Text = "Publishing PDF document(s)...";
                var result = _context.ExportToPdfWithOptions(settings);

                if (result.Success)
                {
                    StatusLabel.Text = "PDF publishing completed successfully!";
                    string msg = result.FilesExported == 1
                        ? $"PDF published successfully:\n{result.ExportedFiles[0]}"
                        : $"Exported {result.FilesExported} PDF files to folder:\n{settings.OutputDirectory}";

                    _context.ShowMessage(msg, "PDF Export Complete");
                    DialogResult = true;
                    Close();
                }
                else
                {
                    StatusLabel.Text = $"Export failed: {result.ErrorMessage}";
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
