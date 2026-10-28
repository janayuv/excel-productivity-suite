using System;
using System.IO;
using System.Text;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Files;
using ExcelProductivitySuite.UI.Controls;
using Microsoft.Win32;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class InsertTextFileDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public InsertTextFileDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            EncodingCombo.Items.Add("UTF-8");
            EncodingCombo.Items.Add("ANSI / Windows-1252");
            EncodingCombo.Items.Add("Unicode (UTF-16)");
            EncodingCombo.Items.Add("ASCII");
            EncodingCombo.SelectedIndex = 0;
        }

        private Encoding SelectedEncoding
        {
            get
            {
                int idx = EncodingCombo.SelectedIndex;
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
                Title = "Select Text or Log File",
                Filter = "Text Files (*.txt;*.log;*.csv;*.json;*.xml)|*.txt;*.log;*.csv;*.json;*.xml|Log Files (*.log)|*.log|All Files (*.*)|*.*",
                CheckFileExists = true
            };

            if (ofd.ShowDialog() == true)
            {
                FilePathText.Text = ofd.FileName;
            }
        }

        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            string path = FilePathText.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                StatusLabel.Text = "Please select a valid file to insert.";
                return;
            }

            try
            {
                var mode = SingleCellRadio.IsChecked == true
                    ? TextFileInsertMode.SingleCell
                    : TextFileInsertMode.RowByRow;

                var matrix = TextFileReader.ReadFileToMatrix(path, mode, SelectedEncoding);
                int rStart = _context.SelectionStartRow;
                int cStart = _context.SelectionStartColumn;

                _context.SetRangeValues(rStart, cStart, matrix);

                int rowsInserted = matrix.GetLength(0);
                StatusLabel.Text = $"Inserted {rowsInserted} row(s).";
                _context.ShowMessage($"Inserted '{Path.GetFileName(path)}' ({rowsInserted} row(s)) at row {rStart}, col {cStart}.", "Insert Text File");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error: {ex.Message}";
                _context.ShowMessage($"Failed to insert file: {ex.Message}", "Insert Error", isError: true);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
