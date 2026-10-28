using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class BatchConvertWorkbooksDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private List<string> _matchingFiles = new();

        public BatchConvertWorkbooksDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SourcePatternCombo.Items.Add("Legacy Excel (*.xls)");
            SourcePatternCombo.Items.Add("Modern Excel (*.xlsx)");
            SourcePatternCombo.Items.Add("All Excel Files (*.xls;*.xlsx)");
            SourcePatternCombo.Items.Add("CSV Files (*.csv)");
            SourcePatternCombo.SelectedIndex = 0;

            TargetFormatCombo.Items.Add("Modern Excel (.xlsx)");
            TargetFormatCombo.Items.Add("PDF Document (.pdf)");
            TargetFormatCombo.Items.Add("CSV Delimited (.csv)");
            TargetFormatCombo.Items.Add("Legacy Excel (.xls)");
            TargetFormatCombo.SelectedIndex = 0;

            string myDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            SourceFolderText.Text = myDocs;
            OutputFolderText.Text = Path.Combine(myDocs, "Converted_Workbooks");

            ScanSourceFiles();
        }

        private string CurrentSearchPattern
        {
            get
            {
                return SourcePatternCombo.SelectedIndex switch
                {
                    1 => "*.xlsx",
                    2 => "*.xls*",
                    3 => "*.csv",
                    _ => "*.xls"
                };
            }
        }

        private string TargetExtension
        {
            get
            {
                return TargetFormatCombo.SelectedIndex switch
                {
                    1 => ".pdf",
                    2 => ".csv",
                    3 => ".xls",
                    _ => ".xlsx"
                };
            }
        }

        private void BrowseSource_Click(object sender, RoutedEventArgs e)
        {
            using var fbd = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select Source Folder",
                SelectedPath = SourceFolderText.Text,
                ShowNewFolderButton = false
            };

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SourceFolderText.Text = fbd.SelectedPath;
            }
        }

        private void BrowseOutput_Click(object sender, RoutedEventArgs e)
        {
            using var fbd = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select Output Folder",
                SelectedPath = OutputFolderText.Text,
                ShowNewFolderButton = true
            };

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                OutputFolderText.Text = fbd.SelectedPath;
            }
        }

        private void SourceFolder_TextChanged(object sender, TextChangedEventArgs e)
        {
            ScanSourceFiles();
        }

        private void Filter_Changed(object sender, SelectionChangedEventArgs e)
        {
            ScanSourceFiles();
        }

        private void ScanSourceFiles()
        {
            _matchingFiles.Clear();
            string folder = SourceFolderText?.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                if (FoundCountLabel != null) FoundCountLabel.Text = "Please select a valid source folder.";
                return;
            }

            try
            {
                string pattern = CurrentSearchPattern;
                var files = Directory.GetFiles(folder, pattern, SearchOption.TopDirectoryOnly);

                // If looking for *.xls only, exclude *.xlsx and *.xlsm
                if (SourcePatternCombo.SelectedIndex == 0)
                {
                    files = files.Where(f => string.Equals(Path.GetExtension(f), ".xls", StringComparison.OrdinalIgnoreCase)).ToArray();
                }

                _matchingFiles.AddRange(files);
                if (FoundCountLabel != null)
                {
                    FoundCountLabel.Text = $"Found {_matchingFiles.Count} matching file(s) in source directory.";
                }
            }
            catch (Exception ex)
            {
                if (FoundCountLabel != null) FoundCountLabel.Text = $"Scan error: {ex.Message}";
            }
        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            if (_matchingFiles.Count == 0)
            {
                StatusLabel.Text = "No files found to convert.";
                return;
            }

            string outDir = OutputFolderText.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(outDir))
            {
                StatusLabel.Text = "Please specify an output directory.";
                return;
            }

            try
            {
                int converted = _context.BatchConvertWorkbooks(_matchingFiles, outDir, TargetExtension);
                StatusLabel.Text = $"Successfully converted {converted} file(s).";
                _context.ShowMessage($"Converted {converted} of {_matchingFiles.Count} file(s) to '{outDir}'.", "Batch Convert Workbooks");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Conversion error: {ex.Message}";
                _context.ShowMessage($"Conversion failed: {ex.Message}", "Conversion Error", isError: true);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
