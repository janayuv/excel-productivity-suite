using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FrequencyDistributionDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;
        private List<FrequencyRecord> _currentRecords = new();

        public FrequencyDistributionDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdateDistribution();
        }

        private FrequencyDistributionOptions GetOptions()
        {
            var options = new FrequencyDistributionOptions();

            options.CaseSensitive = CaseSensitiveCheck?.IsChecked == true;
            options.IncludeBlanks = IncludeBlanksCheck?.IsChecked == true;

            int sortIdx = SortCombo?.SelectedIndex ?? 0;
            options.SortOrder = sortIdx switch
            {
                1 => FrequencySortOrder.FrequencyAscending,
                2 => FrequencySortOrder.ValueAscending,
                3 => FrequencySortOrder.ValueDescending,
                _ => FrequencySortOrder.FrequencyDescending
            };

            return options;
        }

        private void UpdateDistribution()
        {
            if (DistributionGrid == null || StatsSummaryText == null) return;

            if (_matrix == null || _matrix.Length == 0)
            {
                StatsSummaryText.Text = "No range selected or empty cells.";
                DistributionGrid.ItemsSource = null;
                return;
            }

            var options = GetOptions();
            _currentRecords = FrequencyDistributionCalculator.Calculate(_matrix, options);
            DistributionGrid.ItemsSource = _currentRecords;

            int totalSamples = _currentRecords.Sum(r => r.Count);
            StatsSummaryText.Text = $"Distinct Values: {_currentRecords.Count} | Total Samples: {totalSamples}";
        }

        private void Option_Changed(object sender, RoutedEventArgs e)
        {
            UpdateDistribution();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentRecords.Count == 0)
            {
                _context.ShowMessage("No data available to generate frequency distribution.", "Frequency Distribution", isError: true);
                return;
            }

            var outputMatrix = FrequencyDistributionCalculator.GenerateMatrix(_currentRecords, includeHeaders: true);

            if (NewSheetRadio?.IsChecked == true)
            {
                string sheetName = "Freq_Stats_" + DateTime.Now.ToString("HHmmss");
                bool success = _context.WriteMatrixToNewWorksheet(sheetName, outputMatrix);
                if (success)
                {
                    _context.ShowMessage($"Frequency distribution table created successfully in worksheet '{sheetName}'.", "Frequency Distribution", isError: false);
                    Close();
                }
                else
                {
                    _context.ShowMessage("Failed to create frequency distribution worksheet.", "Frequency Distribution", isError: true);
                }
            }
            else
            {
                // Adjacent to selection: Write to active sheet starting at row 1 of next empty column
                int colCount = _matrix?.GetLength(1) ?? 1;
                int startRow = _context.SelectionStartRow > 0 ? _context.SelectionStartRow : 1;
                int targetCol = (_context.SelectionStartColumn > 0 ? _context.SelectionStartColumn : 1) + colCount + 1;

                string targetAddr = $"R{startRow}C{targetCol}";
                bool success = _context.WriteMatrixToAddress(targetAddr, outputMatrix);
                if (success)
                {
                    _context.ShowMessage("Frequency distribution table placed adjacent to active selection.", "Frequency Distribution", isError: false);
                    Close();
                }
                else
                {
                    _context.ShowMessage("Failed to write frequency distribution table adjacent to selection.", "Frequency Distribution", isError: true);
                }
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
