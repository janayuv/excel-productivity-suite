using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class DataEntryFormDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private TableDataModel? _tableModel;
        private readonly List<TextBox> _fieldBoxes = new List<TextBox>();
        private int _currentRecordIndex = 0;

        public DataEntryFormDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadTable();
        }

        private void LoadTable()
        {
            _tableModel = _context.GetActiveTableData();
            if (_tableModel == null || _tableModel.Headers.Count == 0)
            {
                StatusText.Text = "No table or data range detected. Please select a table or range with headers.";
                RecordStatusText.Text = "No Data";
                PrevBtn.IsEnabled = false;
                NextBtn.IsEnabled = false;
                return;
            }

            BuildFields();

            if (_tableModel.Rows.Count > 0)
            {
                _currentRecordIndex = 0;
            }
            else
            {
                _currentRecordIndex = -1; // new record
            }

            DisplayRecord();
        }

        private void BuildFields()
        {
            FieldsContainer.Children.Clear();
            _fieldBoxes.Clear();

            if (_tableModel == null) return;

            foreach (string header in _tableModel.Headers)
            {
                var label = new TextBlock
                {
                    Text = header,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = (System.Windows.Media.Brush)FindResource("TextPrimaryBrush"),
                    Margin = new Thickness(0, 4, 0, 2)
                };

                var textBox = new TextBox
                {
                    Style = (Style)FindResource("ModernTextBoxStyle"),
                    Margin = new Thickness(0, 0, 0, 6)
                };

                FieldsContainer.Children.Add(label);
                FieldsContainer.Children.Add(textBox);
                _fieldBoxes.Add(textBox);
            }
        }

        private void DisplayRecord()
        {
            if (_tableModel == null) return;

            int totalRows = _tableModel.Rows.Count;

            if (_currentRecordIndex == -1 || totalRows == 0)
            {
                RecordStatusText.Text = "New Record (Unsaved)";
                PrevBtn.IsEnabled = totalRows > 0;
                NextBtn.IsEnabled = false;
                foreach (var b in _fieldBoxes) b.Text = "";
            }
            else
            {
                RecordStatusText.Text = $"Record {_currentRecordIndex + 1} of {totalRows}";
                PrevBtn.IsEnabled = _currentRecordIndex > 0;
                NextBtn.IsEnabled = _currentRecordIndex < totalRows - 1;

                var rowData = _tableModel.Rows[_currentRecordIndex];
                for (int i = 0; i < _fieldBoxes.Count; i++)
                {
                    _fieldBoxes[i].Text = (i < rowData.Length && rowData[i] != null) ? rowData[i]!.ToString() : "";
                }
            }

            StatusText.Text = $"{_tableModel.Headers.Count} field(s) mapped from active table.";
        }

        private void Prev_Click(object sender, RoutedEventArgs e)
        {
            if (_tableModel == null || _tableModel.Rows.Count == 0) return;

            if (_currentRecordIndex == -1)
            {
                _currentRecordIndex = _tableModel.Rows.Count - 1;
            }
            else if (_currentRecordIndex > 0)
            {
                _currentRecordIndex--;
            }
            DisplayRecord();
        }

        private void Next_Click(object sender, RoutedEventArgs e)
        {
            if (_tableModel == null) return;

            if (_currentRecordIndex >= 0 && _currentRecordIndex < _tableModel.Rows.Count - 1)
            {
                _currentRecordIndex++;
                DisplayRecord();
            }
        }

        private void NewRecord_Click(object sender, RoutedEventArgs e)
        {
            _currentRecordIndex = -1;
            DisplayRecord();
            if (_fieldBoxes.Count > 0)
            {
                _fieldBoxes[0].Focus();
            }
        }

        private void ClearFields_Click(object sender, RoutedEventArgs e)
        {
            foreach (var b in _fieldBoxes) b.Text = "";
        }

        private void SaveRecord_Click(object sender, RoutedEventArgs e)
        {
            if (_tableModel == null) return;

            var newRowVals = new object?[_fieldBoxes.Count];
            for (int i = 0; i < _fieldBoxes.Count; i++)
            {
                newRowVals[i] = _fieldBoxes[i].Text.Trim();
            }

            bool success = _context.AppendTableRow(newRowVals, _tableModel.StartRow, _tableModel.StartCol, _tableModel.Headers.Count);
            if (success)
            {
                _tableModel.Rows.Add(newRowVals);
                _currentRecordIndex = _tableModel.Rows.Count - 1;
                DisplayRecord();
                _context.ShowMessage("New record successfully saved and appended to worksheet table.", "Record Saved");
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
