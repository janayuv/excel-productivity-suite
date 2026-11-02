using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Printing;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class PrintMultipleWorkbooksDialog : BaseDialogWindow
    {
        private readonly IExcelContext? _context;
        private int _currentStep = 1;
        private readonly ObservableCollection<WorkbookPrintItem> _files = new ObservableCollection<WorkbookPrintItem>();
        private CancellationTokenSource? _cts;

        public PrintMultipleWorkbooksDialog()
        {
            InitializeComponent();
            FilesListView.ItemsSource = _files;
            UpdateStepUI();
        }

        public PrintMultipleWorkbooksDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;
            FilesListView.ItemsSource = _files;

            LoadPrinters();
            UpdateStepUI();
        }

        private void LoadPrinters()
        {
            try
            {
                var printers = _context?.GetInstalledPrinterNames() ?? new List<string>();
                string? defaultPrinter = _context?.GetDefaultPrinterName();

                PrintersComboBox.Items.Clear();
                PrintersComboBox.Items.Add("(System Default Printer)");

                foreach (var p in printers)
                {
                    PrintersComboBox.Items.Add(p);
                }

                if (!string.IsNullOrWhiteSpace(defaultPrinter) && printers.Contains(defaultPrinter))
                {
                    PrintersComboBox.SelectedItem = defaultPrinter;
                }
                else
                {
                    PrintersComboBox.SelectedIndex = 0;
                }
            }
            catch
            {
                PrintersComboBox.Items.Add("(System Default Printer)");
                PrintersComboBox.SelectedIndex = 0;
            }
        }

        private void UpdateStepUI()
        {
            // Panels
            Step1Panel.Visibility = _currentStep == 1 ? Visibility.Visible : Visibility.Collapsed;
            Step2Panel.Visibility = _currentStep == 2 ? Visibility.Visible : Visibility.Collapsed;
            Step3Panel.Visibility = _currentStep == 3 ? Visibility.Visible : Visibility.Collapsed;
            Step4Panel.Visibility = _currentStep == 4 ? Visibility.Visible : Visibility.Collapsed;

            // Badges
            SetBadgeStyle(Step1Badge, _currentStep == 1, _currentStep > 1);
            SetBadgeStyle(Step2Badge, _currentStep == 2, _currentStep > 2);
            SetBadgeStyle(Step3Badge, _currentStep == 3, _currentStep > 3);
            SetBadgeStyle(Step4Badge, _currentStep == 4, false);

            // Nav Buttons
            BackButton.IsEnabled = _currentStep > 1;
            NextButton.Visibility = _currentStep < 4 ? Visibility.Visible : Visibility.Collapsed;
            FinishButton.Visibility = _currentStep == 4 ? Visibility.Visible : Visibility.Collapsed;

            // Footer Text
            switch (_currentStep)
            {
                case 1:
                    StepFooterText.Text = "Step 1 of 4: Specify source folder and file formats";
                    break;
                case 2:
                    StepFooterText.Text = "Step 2 of 4: Select and verify workbooks to print";
                    UpdateFileCountSummary();
                    break;
                case 3:
                    StepFooterText.Text = "Step 3 of 4: Choose sheet scope, page range, and copies";
                    break;
                case 4:
                    StepFooterText.Text = "Step 4 of 4: Select printer and configure background options";
                    UpdateReadySummary();
                    break;
            }
        }

        private void SetBadgeStyle(Border badge, bool isCurrent, bool isCompleted)
        {
            if (isCurrent)
            {
                badge.Background = new SolidColorBrush(Color.FromRgb(239, 246, 255)); // #EFF6FF
                badge.BorderBrush = new SolidColorBrush(Color.FromRgb(59, 130, 246)); // #3B82F6
                if (badge.Child is TextBlock tb)
                {
                    tb.Foreground = new SolidColorBrush(Color.FromRgb(29, 78, 216)); // #1D4ED8
                    tb.FontWeight = FontWeights.Bold;
                }
            }
            else if (isCompleted)
            {
                badge.Background = new SolidColorBrush(Color.FromRgb(240, 253, 244)); // #F0FDF4
                badge.BorderBrush = new SolidColorBrush(Color.FromRgb(34, 197, 94)); // #22C55E
                if (badge.Child is TextBlock tb)
                {
                    tb.Foreground = new SolidColorBrush(Color.FromRgb(21, 128, 61)); // #15803D
                    tb.FontWeight = FontWeights.SemiBold;
                }
            }
            else
            {
                badge.Background = new SolidColorBrush(Color.FromRgb(243, 244, 246)); // #F3F4F6
                badge.BorderBrush = new SolidColorBrush(Color.FromRgb(229, 231, 235)); // #E5E7EB
                if (badge.Child is TextBlock tb)
                {
                    tb.Foreground = new SolidColorBrush(Color.FromRgb(107, 114, 128)); // #6B7280
                    tb.FontWeight = FontWeights.Normal;
                }
            }
        }

        private void Next_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStep == 1)
            {
                string dir = FolderPathBox.Text.Trim();
                if (RbSameDir.IsChecked == true && (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir)))
                {
                    _context?.ShowMessage("Please specify a valid folder directory.", "Directory Required", isError: true);
                    return;
                }

                // If moving from Step 1 to Step 2 and in SameDir mode, automatically scan if files list is empty
                if (RbSameDir.IsChecked == true && _files.Count == 0 && Directory.Exists(dir))
                {
                    ScanDirectoryFiles(dir);
                }
            }
            else if (_currentStep == 2)
            {
                int selectedCount = _files.Count(f => f.IsSelected);
                if (selectedCount == 0)
                {
                    _context?.ShowMessage("Please select at least one workbook to print.", "Selection Required", isError: true);
                    return;
                }
            }
            else if (_currentStep == 3)
            {
                if (RbScopeSpecific.IsChecked == true)
                {
                    string spec = SpecificSheetsBox.Text.Trim();
                    if (string.IsNullOrWhiteSpace(spec))
                    {
                        _context?.ShowMessage("Please enter at least one specific sheet name.", "Sheet Name Required", isError: true);
                        return;
                    }
                }

                if (RbPageCustom.IsChecked == true)
                {
                    if (!int.TryParse(FromPageBox.Text.Trim(), out int from) || from < 1 ||
                        !int.TryParse(ToPageBox.Text.Trim(), out int to) || to < from)
                    {
                        _context?.ShowMessage("Please specify a valid page range (From >= 1 and To >= From).", "Invalid Page Range", isError: true);
                        return;
                    }
                }

                if (!int.TryParse(CopiesBox.Text.Trim(), out int copies) || copies < 1 || copies > 99)
                {
                    _context?.ShowMessage("Number of copies must be between 1 and 99.", "Invalid Copies", isError: true);
                    return;
                }
            }

            if (_currentStep < 4)
            {
                _currentStep++;
                UpdateStepUI();
            }
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            if (_currentStep > 1)
            {
                _currentStep--;
                UpdateStepUI();
            }
        }

        private void FolderMode_Changed(object sender, RoutedEventArgs e)
        {
            if (FolderPathBox == null) return;
            FolderPathBox.IsEnabled = RbSameDir?.IsChecked == true;
        }

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
            {
                fbd.Description = "Select Folder Containing Workbooks to Print";
                if (!string.IsNullOrWhiteSpace(FolderPathBox.Text) && Directory.Exists(FolderPathBox.Text))
                {
                    fbd.SelectedPath = FolderPathBox.Text;
                }

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    FolderPathBox.Text = fbd.SelectedPath;
                    ScanDirectoryFiles(fbd.SelectedPath);
                }
            }
        }

        private void ScanDirectoryFiles(string directory)
        {
            bool rec = SubfoldersCheck.IsChecked == true;
            bool incExcel = CheckExcelFiles.IsChecked == true;
            bool incCsv = CheckCsvFiles.IsChecked == true;
            bool incText = CheckTextFiles.IsChecked == true;
            string? customFilter = CustomFilterBox.Text.Trim();

            var scanned = PrintWorkbooksEngine.ScanFiles(directory, rec, incExcel, incCsv, incText, customFilter);

            _files.Clear();
            foreach (var item in scanned)
            {
                _files.Add(item);
            }

            UpdateFileCountSummary();
        }

        private void AddFiles_Click(object sender, RoutedEventArgs e)
        {
            using (var ofd = new System.Windows.Forms.OpenFileDialog())
            {
                ofd.Title = "Select Workbooks to Print";
                ofd.Filter = "All Supported Files|*.xlsx;*.xls;*.xlsm;*.xlsb;*.csv;*.txt|Excel Workbooks (*.xlsx;*.xls)|*.xlsx;*.xls;*.xlsm;*.xlsb|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
                ofd.Multiselect = true;

                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var existingPaths = new HashSet<string>(_files.Select(f => f.FilePath), StringComparer.OrdinalIgnoreCase);
                    foreach (var path in ofd.FileNames)
                    {
                        if (!existingPaths.Contains(path) && File.Exists(path))
                        {
                            var fi = new FileInfo(path);
                            _files.Add(new WorkbookPrintItem
                            {
                                FilePath = fi.FullName,
                                FileName = fi.Name,
                                Directory = fi.DirectoryName ?? string.Empty,
                                SizeBytes = fi.Length,
                                FileSizeFormatted = WorkbookPrintItem.FormatFileSize(fi.Length),
                                LastModified = fi.LastWriteTime,
                                IsSelected = true
                            });
                        }
                    }
                    UpdateFileCountSummary();
                }
            }
        }

        private void AddFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
            {
                fbd.Description = "Select Additional Folder to Scan for Workbooks";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    bool rec = SubfoldersCheck.IsChecked == true;
                    bool incExcel = CheckExcelFiles.IsChecked == true;
                    bool incCsv = CheckCsvFiles.IsChecked == true;
                    bool incText = CheckTextFiles.IsChecked == true;
                    string? customFilter = CustomFilterBox.Text.Trim();

                    var scanned = PrintWorkbooksEngine.ScanFiles(fbd.SelectedPath, rec, incExcel, incCsv, incText, customFilter);
                    var existingPaths = new HashSet<string>(_files.Select(f => f.FilePath), StringComparer.OrdinalIgnoreCase);

                    foreach (var item in scanned)
                    {
                        if (!existingPaths.Contains(item.FilePath))
                        {
                            _files.Add(item);
                        }
                    }
                    UpdateFileCountSummary();
                }
            }
        }

        private void RemoveFiles_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = FilesListView.SelectedItems.Cast<WorkbookPrintItem>().ToList();
            foreach (var item in selectedItems)
            {
                _files.Remove(item);
            }
            UpdateFileCountSummary();
        }

        private void ClearFiles_Click(object sender, RoutedEventArgs e)
        {
            _files.Clear();
            UpdateFileCountSummary();
        }

        private void ScanRefresh_Click(object sender, RoutedEventArgs e)
        {
            string dir = FolderPathBox.Text.Trim();
            if (Directory.Exists(dir))
            {
                ScanDirectoryFiles(dir);
            }
            else
            {
                _context?.ShowMessage("Please specify a valid folder path in Step 1 to scan.", "Directory Not Found");
            }
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var f in _files) f.IsSelected = true;
            FilesListView.Items.Refresh();
            UpdateFileCountSummary();
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var f in _files) f.IsSelected = false;
            FilesListView.Items.Refresh();
            UpdateFileCountSummary();
        }

        private void UpdateFileCountSummary()
        {
            if (FileCountSummaryText == null) return;
            int total = _files.Count;
            int selected = _files.Count(f => f.IsSelected);
            FileCountSummaryText.Text = $"{total} file(s) found, {selected} selected";
        }

        private void Scope_Changed(object sender, RoutedEventArgs e)
        {
            if (SpecificSheetsPanel == null) return;
            SpecificSheetsPanel.IsEnabled = RbScopeSpecific.IsChecked == true;
        }

        private void PageRange_Changed(object sender, RoutedEventArgs e)
        {
            if (FromPageBox == null || ToPageBox == null) return;
            bool custom = RbPageCustom.IsChecked == true;
            FromPageBox.IsEnabled = custom;
            ToPageBox.IsEnabled = custom;
        }

        private void UpdateReadySummary()
        {
            if (ReadySummaryText == null) return;
            int selected = _files.Count(f => f.IsSelected);
            string printer = PrintersComboBox.SelectedItem?.ToString() ?? "Default Printer";
            string scope = "Entire Workbook";
            if (RbScopeWorksheetsOnly.IsChecked == true) scope = "Worksheets Only";
            else if (RbScopeChartSheetsOnly.IsChecked == true) scope = "Chart Sheets Only";
            else if (RbScopeWorksheetsAndCharts.IsChecked == true) scope = "Worksheets & Charts";
            else if (RbScopeSpecific.IsChecked == true) scope = $"Specific Sheets ({SpecificSheetsBox.Text.Trim()})";

            ReadySummaryText.Text = $"Ready to print {selected} selected workbook(s).\n" +
                                   $"Scope: {scope} | Printer: {printer} | Copies: {CopiesBox.Text.Trim()}";
        }

        private async void Finish_Click(object sender, RoutedEventArgs e)
        {
            var config = BuildConfig();
            var (isValid, errorMsg) = PrintWorkbooksEngine.ValidateConfig(config);
            if (!isValid)
            {
                _context?.ShowMessage(errorMsg ?? "Invalid configuration.", "Validation Error", isError: true);
                return;
            }

            // Disable UI controls while printing
            FinishButton.IsEnabled = false;
            BackButton.IsEnabled = false;
            BatchProgressBar.Visibility = Visibility.Visible;
            StatusProgressText.Visibility = Visibility.Visible;
            BatchProgressBar.Value = 0;
            StatusProgressText.Text = "Initializing batch print operation...";

            _cts = new CancellationTokenSource();
            var progress = new Progress<(int current, int total, string currentFile)>(info =>
            {
                int pct = (int)((double)info.current / Math.Max(1, info.total) * 100);
                BatchProgressBar.Value = pct;
                StatusProgressText.Text = $"Printing {info.current} of {info.total}: {info.currentFile}...";
            });

            PrintWorkbooksResult? result = null;
            try
            {
                if (_context != null)
                {
                    result = await Task.Run(() => _context.BatchPrintWorkbooks(config, progress, _cts.Token));
                }
            }
            catch (Exception ex)
            {
                _context?.ShowMessage($"Error during batch print: {ex.Message}", "Print Error", isError: true);
            }
            finally
            {
                FinishButton.IsEnabled = true;
                BackButton.IsEnabled = true;
                BatchProgressBar.Visibility = Visibility.Collapsed;
                StatusProgressText.Visibility = Visibility.Collapsed;
            }

            if (result != null)
            {
                string summaryMsg = $"Batch printing completed in {result.Duration.TotalSeconds:F1}s!\n\n" +
                                   $"• Total Workbooks: {result.TotalWorkbooks}\n" +
                                   $"• Printed Workbooks: {result.PrintedWorkbooks}\n" +
                                   $"• Total Sheets Printed: {result.TotalSheetsPrinted}\n" +
                                   $"• Failed: {result.FailedWorkbooks}";

                if (result.FailedWorkbooks > 0)
                {
                    summaryMsg += $"\n\nErrors:\n" + string.Join("\n", result.Errors.Take(3));
                }

                _context?.ShowMessage(summaryMsg, "Print Multiple Workbooks Complete");
                DialogResult = true;
                Close();
            }
        }

        private PrintWorkbooksConfig BuildConfig()
        {
            var config = new PrintWorkbooksConfig
            {
                FolderMode = RbSameDir.IsChecked == true ? PrintFolderMode.SameDirectory : PrintFolderMode.MultipleDirectories,
                DirectoryPath = FolderPathBox.Text.Trim(),
                IncludeSubdirectories = SubfoldersCheck.IsChecked == true,
                SelectedFiles = _files.Where(f => f.IsSelected).ToList(),
                IncludeExcelFiles = CheckExcelFiles.IsChecked == true,
                IncludeCsvFiles = CheckCsvFiles.IsChecked == true,
                IncludeTextFiles = CheckTextFiles.IsChecked == true,
                CustomExtensionFilter = CustomFilterBox.Text.Trim(),
                QuietMode = CheckQuietMode.IsChecked == true,
                DisableMacros = CheckDisableMacros.IsChecked == true,
                DoNotUpdateLinks = CheckDoNotUpdateLinks.IsChecked == true,
                GenerateSummaryReport = CheckGenerateReport.IsChecked == true
            };

            // Sheet Scope
            if (RbScopeWorksheetsOnly.IsChecked == true) config.SheetScope = PrintSheetScope.WorksheetsOnly;
            else if (RbScopeChartSheetsOnly.IsChecked == true) config.SheetScope = PrintSheetScope.ChartSheetsOnly;
            else if (RbScopeWorksheetsAndCharts.IsChecked == true) config.SheetScope = PrintSheetScope.WorksheetsAndCharts;
            else if (RbScopeSpecific.IsChecked == true)
            {
                config.SheetScope = PrintSheetScope.SpecificSheets;
                config.SpecificSheetNames = SpecificSheetsBox.Text
                    .Split(new[] { ',', ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => s.Length > 0)
                    .ToList();
            }
            else
            {
                config.SheetScope = PrintSheetScope.EntireWorkbook;
            }

            // Page range
            if (RbPageFirst.IsChecked == true)
            {
                config.PageRangeMode = PrintPageRangeMode.FirstPageOnly;
            }
            else if (RbPageCustom.IsChecked == true)
            {
                config.PageRangeMode = PrintPageRangeMode.CustomRange;
                int.TryParse(FromPageBox.Text.Trim(), out int from);
                int.TryParse(ToPageBox.Text.Trim(), out int to);
                config.FromPage = Math.Max(1, from);
                config.ToPage = Math.Max(config.FromPage, to);
            }
            else
            {
                config.PageRangeMode = PrintPageRangeMode.AllPages;
            }

            // Copies
            if (int.TryParse(CopiesBox.Text.Trim(), out int copies))
            {
                config.Copies = Math.Max(1, Math.Min(99, copies));
            }

            // Printer
            string selectedPrinter = PrintersComboBox.SelectedItem?.ToString() ?? "";
            if (!string.IsNullOrWhiteSpace(selectedPrinter) && !selectedPrinter.StartsWith("("))
            {
                config.PrinterName = selectedPrinter;
            }

            return config;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
            DialogResult = false;
            Close();
        }
    }
}
