using System;
using System.Text;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Analysis;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class QuickStatisticsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly StatisticsResult _stats;
        private readonly string? _address;

        public QuickStatisticsDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            _address = _context.SelectionAddress ?? "Selection";
            ScopeSubtitle.Text = $"Source Range: {_address} ({_context.ActiveWorksheetName ?? "Active Sheet"})";

            var values = _context.GetSelectedValues();
            _stats = QuickStatisticsEngine.Calculate(values);

            PopulateUI();
        }

        private void PopulateUI()
        {
            // Cards
            CardSumText.Text = _stats.NumericCells > 0 ? FormatMetricNumber(_stats.Sum) : "0";
            CardCountSubtext.Text = $"{_stats.NumericCells:N0} numeric cells ({_stats.TotalCells:N0} total)";

            CardMeanText.Text = _stats.NumericCells > 0 ? FormatMetricNumber(_stats.Mean) : "0";
            CardMedianSubtext.Text = $"Median: {(_stats.NumericCells > 0 ? FormatMetricNumber(_stats.Median) : "0")}";

            CardMinMaxText.Text = _stats.NumericCells > 0 ? $"{FormatMetricNumber(_stats.Min)} / {FormatMetricNumber(_stats.Max)}" : "0 / 0";
            CardRangeSubtext.Text = $"Range: {(_stats.NumericCells > 0 ? FormatMetricNumber(_stats.Range) : "0")}";

            CardStdDevText.Text = _stats.NumericCells > 0 ? FormatMetricNumber(_stats.SampleStdDev) : "0";
            CardVarianceSubtext.Text = $"Variance: {(_stats.NumericCells > 0 ? FormatMetricNumber(_stats.SampleVariance) : "0")}";

            MetricsDataGrid.ItemsSource = _stats.MetricItems;
        }

        private string FormatMetricNumber(double val)
        {
            if (Math.Abs(val) >= 1_000_000 || (Math.Abs(val) > 0 && Math.Abs(val) < 0.0001))
            {
                return val.ToString("G6");
            }
            if (Math.Abs(val % 1) < double.Epsilon)
            {
                return val.ToString("N0");
            }
            return val.ToString("N4").TrimEnd('0').TrimEnd('.');
        }

        private void CopyAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Category\tStatistic\tValue");
                foreach (var item in _stats.MetricItems)
                {
                    sb.AppendLine($"{item.Category}\t{item.Name}\t{item.FormattedValue}");
                }
                Clipboard.SetText(sb.ToString());
                _context.ShowMessage("All statistical measures copied to clipboard.", "Clipboard");
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Could not copy to clipboard: {ex.Message}", "Clipboard Error", isError: true);
            }
        }

        private void CopySelected_Click(object sender, RoutedEventArgs e)
        {
            if (MetricsDataGrid.SelectedItem is MetricItem selected)
            {
                try
                {
                    Clipboard.SetText($"{selected.Name}: {selected.FormattedValue}");
                    _context.ShowMessage($"Copied '{selected.Name}' to clipboard.", "Clipboard");
                }
                catch (Exception ex)
                {
                    _context.ShowMessage($"Could not copy: {ex.Message}", "Clipboard Error", isError: true);
                }
            }
            else
            {
                _context.ShowMessage("Please select a metric row to copy.", "Selection Required");
            }
        }

        private void InsertSheet_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                object[,] summaryTable = QuickStatisticsEngine.GenerateSummaryTable(_stats, _address);
                string sheetName = "QuickStats_" + DateTime.Now.ToString("HHmmss");

                bool success = _context.CreateWorksheetWithData(sheetName, summaryTable, activate: true);
                if (success)
                {
                    _context.ShowMessage($"Statistics summary table inserted into new sheet '{sheetName}'.", "Sheet Created");
                    DialogResult = true;
                    Close();
                }
                else
                {
                    _context.ShowMessage("Failed to create new worksheet for statistics.", "Error", isError: true);
                }
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Error creating summary sheet: {ex.Message}", "Error", isError: true);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
