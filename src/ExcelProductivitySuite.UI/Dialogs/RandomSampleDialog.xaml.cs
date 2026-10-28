using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class RandomSampleDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public RandomSampleDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            UpdateStatus();
        }

        private void UpdateStatus()
        {
            int totalCells = _context.SelectedCellCount;
            int rows = _context.SelectionRowCount;
            StatusText.Text = $"Active selection covers {rows} row(s) ({totalCells} cells).";
        }

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;
            bool byCount = ByCountRadio.IsChecked == true;
            CountInputBox.IsEnabled = byCount;
            PercentInputBox.IsEnabled = !byCount;
        }

        private void SampleBtn_Click(object sender, RoutedEventArgs e)
        {
            object?[,]? sourceMatrix = _context.GetSelectedValues();
            if (sourceMatrix == null || sourceMatrix.Length == 0)
            {
                _context.ShowMessage("Please select a valid dataset range in Excel to sample.", "Selection Required", isError: true);
                return;
            }

            var options = new RandomSampleOptions
            {
                Mode = ByCountRadio.IsChecked == true ? RandomSampleMode.FixedCount : RandomSampleMode.Percentage,
                HasHeaders = HasHeadersCheck.IsChecked == true,
                WithReplacement = WithReplacementCheck.IsChecked == true
            };

            if (options.Mode == RandomSampleMode.FixedCount)
            {
                if (!int.TryParse(CountInputBox.Text, out int c) || c < 1)
                {
                    _context.ShowMessage("Please specify a valid row sample count (integer >= 1).", "Invalid Input", isError: true);
                    return;
                }
                options.Count = c;
            }
            else
            {
                if (!double.TryParse(PercentInputBox.Text, out double p) || p <= 0 || p > 100)
                {
                    _context.ShowMessage("Please specify a valid sampling percentage between 0.1% and 100%.", "Invalid Input", isError: true);
                    return;
                }
                options.Percentage = p;
            }

            try
            {
                var result = RandomSampler.ExtractSample(sourceMatrix, options);
                if (result.SampleMatrix == null || result.SampleMatrix.Length == 0)
                {
                    _context.ShowMessage("Failed to extract sample: no rows selected.", "Sampling Error", isError: true);
                    return;
                }

                if (DestNewSheetRadio.IsChecked == true)
                {
                    _context.WriteMatrixToNewWorksheet("Random_Sample", result.SampleMatrix);
                    _context.ShowMessage($"Extracted {result.SampledRowsCount} sample rows (from {result.TotalDataRows} total rows) to new worksheet 'Random_Sample'.", "Sample Complete");
                }
                else
                {
                    string target = DestCellBox.Text?.Trim() ?? "$G$1";
                    _context.WriteMatrixToAddress(target, result.SampleMatrix);
                    _context.ShowMessage($"Extracted {result.SampledRowsCount} sample rows to {target}.", "Sample Complete");
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Sampling failed: {ex.Message}", "Sampling Error", isError: true);
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
