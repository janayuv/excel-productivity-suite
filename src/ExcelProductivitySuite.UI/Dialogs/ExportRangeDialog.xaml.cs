using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Files;
using ExcelProductivitySuite.UI.Controls;
using Microsoft.Win32;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ExportRangeDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public ExportRangeDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeAddressLabel.Text = _context.SelectionAddress ?? "Selection";

            FormatCombo.Items.Add("CSV (Comma Separated)");
            FormatCombo.Items.Add("TSV (Tab Separated)");
            FormatCombo.Items.Add("Custom Delimited");
            FormatCombo.Items.Add("JSON (Structured Array)");
            FormatCombo.SelectedIndex = 0;

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            DestinationPathText.Text = Path.Combine(desktop, "Exported_Range.csv");
        }

        private void Format_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CustomDelimLabel == null || CustomDelimText == null || DestinationPathText == null) return;

            bool isCustom = FormatCombo.SelectedIndex == 2;
            CustomDelimLabel.Visibility = isCustom ? Visibility.Visible : Visibility.Collapsed;
            CustomDelimText.Visibility = isCustom ? Visibility.Visible : Visibility.Collapsed;

            string currentPath = DestinationPathText.Text;
            if (!string.IsNullOrWhiteSpace(currentPath))
            {
                string dir = Path.GetDirectoryName(currentPath) ?? "";
                string baseName = Path.GetFileNameWithoutExtension(currentPath);
                string newExt = FormatCombo.SelectedIndex switch
                {
                    1 => ".tsv",
                    2 => ".txt",
                    3 => ".json",
                    _ => ".csv"
                };
                DestinationPathText.Text = Path.Combine(dir, baseName + newExt);
            }
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            string filter = FormatCombo.SelectedIndex switch
            {
                1 => "TSV Files (*.tsv)|*.tsv|All Files (*.*)|*.*",
                2 => "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                3 => "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                _ => "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            };

            var sfd = new SaveFileDialog
            {
                Title = "Save Exported Range As",
                Filter = filter,
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
                StatusLabel.Text = "Please enter a valid destination file path.";
                return;
            }

            try
            {
                object?[,]? matrix = _context.GetSelectedValues();
                if (matrix == null || matrix.Length == 0)
                {
                    StatusLabel.Text = "No range selected to export.";
                    return;
                }

                var options = new RangeExportOptions
                {
                    Format = FormatCombo.SelectedIndex switch
                    {
                        1 => ExportFileFormat.Tsv,
                        2 => ExportFileFormat.CustomDelimited,
                        3 => ExportFileFormat.Json,
                        _ => ExportFileFormat.Csv
                    },
                    Delimiter = CustomDelimText.Text,
                    HasHeaders = HasHeadersCheck.IsChecked == true,
                    QuoteAllFields = QuoteAllCheck.IsChecked == true
                };

                RangeExporter.ExportToFile(matrix, dest, options);
                StatusLabel.Text = "Export successful!";
                _context.ShowMessage($"Exported range to '{dest}' successfully.", "Export Range");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Export failed: {ex.Message}";
                _context.ShowMessage($"Failed to export range: {ex.Message}", "Export Error", isError: true);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
