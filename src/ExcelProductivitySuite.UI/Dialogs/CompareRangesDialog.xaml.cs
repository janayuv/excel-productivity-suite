using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class CompareRangesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _sourceMatrix;

        public CompareRangesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SourceRangeTextBox.Text = _context.SelectionAddress ?? "$A$1";
            _sourceMatrix = _context.GetSelectedValues();

            // Default suggestion for target range
            if (!string.IsNullOrEmpty(_context.SelectionAddress))
            {
                TargetRangeTextBox.Text = SuggestAdjacentTargetAddress(_context);
            }
        }

        private static string SuggestAdjacentTargetAddress(IExcelContext context)
        {
            try
            {
                int rStart = context.SelectionStartRow;
                int rCount = context.SelectionRowCount;
                int cStart = context.SelectionStartColumn + context.SelectionColumnCount;
                int cCount = context.SelectionColumnCount;

                string colStartLetter = GetColumnLetter(cStart);
                string colEndLetter = GetColumnLetter(cStart + cCount - 1);
                int rEnd = rStart + rCount - 1;

                if (colStartLetter == colEndLetter && rStart == rEnd)
                {
                    return $"${colStartLetter}${rStart}";
                }
                return $"${colStartLetter}${rStart}:${colEndLetter}${rEnd}";
            }
            catch
            {
                return "";
            }
        }

        private static string GetColumnLetter(int columnNumber)
        {
            string colLetter = "";
            while (columnNumber > 0)
            {
                int mod = (columnNumber - 1) % 26;
                colLetter = (char)('A' + mod) + colLetter;
                columnNumber = (columnNumber - mod) / 26;
            }
            return string.IsNullOrEmpty(colLetter) ? "A" : colLetter;
        }

        private RangeCompareOptions GetOptions()
        {
            var options = new RangeCompareOptions
            {
                MatchCase = MatchCaseCheckBox.IsChecked == true,
                TrimWhitespace = TrimWhitespaceCheckBox.IsChecked == true,
                IgnoreEmptyCells = IgnoreEmptyCheckBox.IsChecked == true
            };

            if (ModeCellByCellRadio.IsChecked == true) options.Mode = RangeCompareMode.CellByCell;
            else if (ModeMissingInTargetRadio.IsChecked == true) options.Mode = RangeCompareMode.FindValuesNotInTarget;
            else if (ModeCommonRadio.IsChecked == true) options.Mode = RangeCompareMode.FindCommonValues;

            return options;
        }

        private void TargetRange_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            SummaryText.Text = "Target range changed. Click 'Run Comparison' to update results.";
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            if (DiscrepancyGrid != null && DiscrepancyGrid.ItemsSource != null)
            {
                RunComparison();
            }
        }

        private void CompareButton_Click(object sender, RoutedEventArgs e)
        {
            RunComparison();
        }

        private void RunComparison()
        {
            if (_sourceMatrix == null)
            {
                _context.ShowMessage("No source cells are selected.", "Compare Ranges", isError: true);
                return;
            }

            string targetAddress = TargetRangeTextBox.Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(targetAddress))
            {
                _context.ShowMessage("Please specify a target range address to compare against.", "Compare Ranges", isError: true);
                return;
            }

            var targetMatrix = _context.GetRangeValuesByAddress(targetAddress);
            if (targetMatrix == null)
            {
                return; // Error already shown in adapter
            }

            var options = GetOptions();
            var result = RangeComparer.Compare(_sourceMatrix, targetMatrix, options);

            DiscrepancyGrid.ItemsSource = result.Differences;

            if (options.Mode == RangeCompareMode.CellByCell)
            {
                SummaryText.Text = $"Compared {result.TotalCellsCompared} cells. Found {result.TotalDifferences} discrepancies.";
            }
            else if (options.Mode == RangeCompareMode.FindValuesNotInTarget)
            {
                SummaryText.Text = $"Found {result.MissingValues.Count} values in source range that do not exist in target.";
            }
            else if (options.Mode == RangeCompareMode.FindCommonValues)
            {
                SummaryText.Text = $"Found {result.CommonValues.Count} matching common values across both ranges.";
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
