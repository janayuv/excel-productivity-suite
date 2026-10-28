using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Analysis;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class TargetVsActualChartDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _data;
        private readonly int _startRow;
        private readonly int _startCol;
        private readonly int _rowCount;
        private readonly int _colCount;

        private class ColumnOption
        {
            public int Index { get; }
            public string DisplayName { get; }

            public ColumnOption(int index, string displayName)
            {
                Index = index;
                DisplayName = displayName;
            }

            public override string ToString() => DisplayName;
        }

        public TargetVsActualChartDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            _startRow = _context.SelectionStartRow;
            _startCol = _context.SelectionStartColumn;
            _rowCount = _context.SelectionRowCount;
            _colCount = _context.SelectionColumnCount;

            RangeAddressText.Text = _context.SelectionAddress ?? $"R{_startRow}C{_startCol}:R{_startRow + _rowCount - 1}C{_startCol + _colCount - 1}";

            _data = _context.GetSelectedValues();

            PopulateColumns();
        }

        private void PopulateColumns()
        {
            if (_colCount < 2)
            {
                StatusText.Text = "Please select a range containing at least 2 columns (e.g. Category, Actual, Target).";
                return;
            }

            var options = new List<ColumnOption>();
            for (int c = 1; c <= _colCount; c++)
            {
                string headerName = $"Column {c}";
                if (_data != null)
                {
                    int rStart = _data.GetLowerBound(0);
                    int cStart = _data.GetLowerBound(1);
                    if (rStart <= _data.GetUpperBound(0) && cStart + c - 1 <= _data.GetUpperBound(1))
                    {
                        var hVal = _data[rStart, cStart + c - 1]?.ToString()?.Trim();
                        if (!string.IsNullOrWhiteSpace(hVal))
                        {
                            headerName = $"Col {c}: {hVal}";
                        }
                    }
                }
                options.Add(new ColumnOption(c, headerName));
            }

            CategoryCombo.ItemsSource = options;
            ActualCombo.ItemsSource = new List<ColumnOption>(options);
            TargetCombo.ItemsSource = new List<ColumnOption>(options);

            // Auto detect
            var detected = TargetVsActualDetector.Detect(_data, _colCount);

            int catIdx = Math.Min(detected.CategoryColumnIndex - 1, options.Count - 1);
            int actIdx = Math.Min(detected.ActualColumnIndex - 1, options.Count - 1);
            int tgtIdx = Math.Min(detected.TargetColumnIndex - 1, options.Count - 1);

            CategoryCombo.SelectedIndex = Math.Max(0, catIdx);
            ActualCombo.SelectedIndex = Math.Max(0, actIdx);
            TargetCombo.SelectedIndex = Math.Max(0, tgtIdx);

            UpdateSummary(detected);
        }

        private void Mapping_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (CategoryCombo.SelectedItem is ColumnOption cat &&
                ActualCombo.SelectedItem is ColumnOption act &&
                TargetCombo.SelectedItem is ColumnOption tgt)
            {
                // Recompute quick summary
                var cfg = new TargetVsActualConfig
                {
                    CategoryColumnIndex = cat.Index,
                    ActualColumnIndex = act.Index,
                    TargetColumnIndex = tgt.Index
                };
                UpdateSummary(cfg);
            }
        }

        private void UpdateSummary(TargetVsActualConfig cfg)
        {
            int rCount = Math.Max(0, _rowCount - 1); // exclude header
            PreviewRowsText.Text = $"{rCount} rows";

            if (cfg.ValidDataRowCount > 0)
            {
                PreviewActualText.Text = cfg.TotalActual.ToString("N1");
                PreviewTargetText.Text = cfg.TotalTarget.ToString("N1");
                string sign = cfg.TotalVariance >= 0 ? "+" : "";
                PreviewVarianceText.Text = $"Net Variance: {sign}{cfg.TotalVariance:N1} ({sign}{cfg.VariancePercent:F1}%)";
            }
            else
            {
                PreviewActualText.Text = "-";
                PreviewTargetText.Text = "-";
                PreviewVarianceText.Text = "Variance: N/A";
            }
        }

        private void CreateChart_Click(object sender, RoutedEventArgs e)
        {
            if (CategoryCombo.SelectedItem is not ColumnOption cat ||
                ActualCombo.SelectedItem is not ColumnOption act ||
                TargetCombo.SelectedItem is not ColumnOption tgt)
            {
                _context.ShowMessage("Please map all three columns (Category, Actual, Target).", "Mapping Required", isError: true);
                return;
            }

            string title = TitleText.Text.Trim();
            if (string.IsNullOrWhiteSpace(title))
            {
                title = "Target vs. Actual Performance";
            }

            bool success = _context.CreateTargetVsActualChart(title, _startRow, _startCol, _rowCount, _colCount, cat.Index, act.Index, tgt.Index);
            if (success)
            {
                _context.ShowMessage("Target vs. Actual chart created successfully on the active worksheet.", "Chart Created");
                DialogResult = true;
                Close();
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
