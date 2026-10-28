using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Files;
using ExcelProductivitySuite.UI.Controls;
using Microsoft.Win32;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ImportCsvDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public ImportCsvDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            PopulateCombos();
        }

        private void PopulateCombos()
        {
            DelimiterCombo.Items.Add("Comma (,)");
            DelimiterCombo.Items.Add("Tab (\\t)");
            DelimiterCombo.Items.Add("Semicolon (;)");
            DelimiterCombo.Items.Add("Pipe (|)");
            DelimiterCombo.Items.Add("Space ( )");
            DelimiterCombo.SelectedIndex = 0;

            EncodingCombo.Items.Add("UTF-8");
            EncodingCombo.Items.Add("ANSI / Windows-1252");
            EncodingCombo.Items.Add("Unicode (UTF-16)");
            EncodingCombo.Items.Add("ASCII");
            EncodingCombo.SelectedIndex = 0;
        }

        private string SelectedDelimiter
        {
            get
            {
                int idx = DelimiterCombo?.SelectedIndex ?? 0;
                return idx switch
                {
                    1 => "\t",
                    2 => ";",
                    3 => "|",
                    4 => " ",
                    _ => ","
                };
            }
        }

        private Encoding SelectedEncoding
        {
            get
            {
                int idx = EncodingCombo?.SelectedIndex ?? 0;
                return idx switch
                {
                    1 => Encoding.Default,
                    2 => Encoding.Unicode,
                    3 => Encoding.ASCII,
                    _ => Encoding.UTF8
                };
            }
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Title = "Select CSV or Delimited File",
                Filter = "Delimited Files (*.csv;*.txt;*.tsv)|*.csv;*.txt;*.tsv|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                CheckFileExists = true
            };

            if (ofd.ShowDialog() == true)
            {
                FilePathText.Text = ofd.FileName;
            }
        }

        private void FilePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void Setting_Changed(object sender, RoutedEventArgs e)
        {
            if (IsLoaded)
            {
                UpdatePreview();
            }
        }

        private void UpdatePreview()
        {
            string path = FilePathText?.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                if (PreviewGrid != null) PreviewGrid.ItemsSource = null;
                if (StatusLabel != null) StatusLabel.Text = "Please select a valid file.";
                return;
            }

            try
            {
                var options = new CsvImportOptions
                {
                    Delimiter = SelectedDelimiter,
                    FileEncoding = SelectedEncoding,
                    TrimWhitespace = TrimWhitespaceCheck?.IsChecked == true,
                    ParseNumbers = ParseNumbersCheck?.IsChecked == true
                };

                var rows = CsvTextImporter.PreviewFirstRows(path, options, 5);
                if (rows.Count == 0)
                {
                    PreviewGrid.ItemsSource = null;
                    StatusLabel.Text = "File is empty.";
                    return;
                }

                var dt = new DataTable();
                int maxCols = 0;
                foreach (var r in rows) if (r.Count > maxCols) maxCols = r.Count;

                for (int c = 0; c < maxCols; c++)
                {
                    string colHeader = rows[0].Count > c && !string.IsNullOrWhiteSpace(rows[0][c])
                        ? rows[0][c]
                        : $"Col {c + 1}";
                    dt.Columns.Add(colHeader);
                }

                for (int r = 1; r < rows.Count; r++)
                {
                    var dr = dt.NewRow();
                    for (int c = 0; c < maxCols; c++)
                    {
                        dr[c] = c < rows[r].Count ? rows[r][c] : "";
                    }
                    dt.Rows.Add(dr);
                }

                PreviewGrid.ItemsSource = dt.DefaultView;
                StatusLabel.Text = $"Previewing {rows.Count} rows across {maxCols} columns.";
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Preview error: {ex.Message}";
            }
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            string path = FilePathText?.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                StatusLabel.Text = "Please select a valid CSV or text file.";
                return;
            }

            try
            {
                var options = new CsvImportOptions
                {
                    Delimiter = SelectedDelimiter,
                    FileEncoding = SelectedEncoding,
                    TrimWhitespace = TrimWhitespaceCheck?.IsChecked == true,
                    ParseNumbers = ParseNumbersCheck?.IsChecked == true
                };

                var matrix = CsvTextImporter.ParseFile(path, options);
                int rows = matrix.GetLength(0);
                int cols = matrix.GetLength(1);

                if (rows == 0 || cols == 0)
                {
                    StatusLabel.Text = "The file contains no data.";
                    return;
                }

                if (NewSheetRadio.IsChecked == true)
                {
                    string sheetName = NewSheetNameText.Text?.Trim() ?? "Imported_Data";
                    _context.CreateWorksheetWithData(sheetName, matrix);
                    _context.ShowMessage($"Imported {rows} rows into new sheet '{sheetName}'.", "Import Complete");
                }
                else
                {
                    int startRow = _context.SelectionStartRow;
                    int startCol = _context.SelectionStartColumn;
                    _context.SetRangeValues(startRow, startCol, matrix);
                    _context.ShowMessage($"Imported {rows} rows starting at row {startRow}, col {startCol}.", "Import Complete");
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Import failed: {ex.Message}";
                _context.ShowMessage($"Failed to import file: {ex.Message}", "Import Error", isError: true);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
