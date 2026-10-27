using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class RecordDisplayModel
    {
        public int RowIndex { get; set; }
        public string CompositeKey { get; set; } = "";
        public string Category { get; set; } = "";
        public string ValuesDisplay { get; set; } = "";
    }

    public partial class FindMissingRecordsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _sourceMatrix;
        private RecordComparisonResult? _lastResult;

        public FindMissingRecordsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SourceRangeTextBox.Text = _context.SelectionAddress ?? "$A$1";
            _sourceMatrix = _context.GetSelectedValues();
        }

        private void CompareButton_Click(object sender, RoutedEventArgs e)
        {
            if (_sourceMatrix == null)
            {
                _context.ShowMessage("No source cells selected.", "Find Missing Records", isError: true);
                return;
            }

            string targetAddress = TargetRangeTextBox.Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(targetAddress))
            {
                _context.ShowMessage("Please enter a valid target range address to compare against.", "Find Missing Records", isError: true);
                return;
            }

            var targetMatrix = _context.GetRangeValuesByAddress(targetAddress);
            if (targetMatrix == null)
            {
                return; // Error already shown in adapter
            }

            int.TryParse(SourceKeyColTextBox.Text, out int sKey);
            int.TryParse(TargetKeyColTextBox.Text, out int tKey);
            sKey = Math.Max(1, sKey) - 1; // 0-based
            tKey = Math.Max(1, tKey) - 1; // 0-based

            var options = new RecordCompareOptions
            {
                HasHeaders = HasHeadersCheckBox.IsChecked == true,
                MatchCase = MatchCaseCheckBox.IsChecked == true,
                TrimWhitespace = TrimSpacesCheckBox.IsChecked == true,
                SourceKeyColumns = new[] { sKey },
                TargetKeyColumns = new[] { tKey }
            };

            _lastResult = RangeRecordComparer.CompareRecords(_sourceMatrix, targetMatrix, options);

            UpdateView();

            if (ExportButton != null) ExportButton.IsEnabled = true;

            SummaryText.Text = $"Source: {_lastResult.TotalSourceRows} rows, Target: {_lastResult.TotalTargetRows} rows. Source-Only: {_lastResult.SourceOnly.Count}, Target-Only: {_lastResult.TargetOnly.Count}, Matching: {_lastResult.Matching.Count}.";
        }

        private void CategoryChanged(object sender, RoutedEventArgs e)
        {
            UpdateView();
        }

        private void UpdateView()
        {
            if (_lastResult == null || ResultsDataGrid == null) return;

            RecordResultCategory category = GetSelectedCategory();
            List<RecordMatchItem> list = category switch
            {
                RecordResultCategory.SourceOnly => _lastResult.SourceOnly,
                RecordResultCategory.TargetOnly => _lastResult.TargetOnly,
                RecordResultCategory.Matching => _lastResult.Matching,
                RecordResultCategory.SourceDuplicates => _lastResult.SourceDuplicates,
                RecordResultCategory.TargetDuplicates => _lastResult.TargetDuplicates,
                _ => _lastResult.SourceOnly
            };

            var models = list.Select(item => new RecordDisplayModel
            {
                RowIndex = item.RowIndex,
                CompositeKey = item.CompositeKey,
                Category = item.Category.ToString(),
                ValuesDisplay = string.Join(" | ", item.Values.Select(v => v?.ToString() ?? ""))
            }).ToList();

            ResultsDataGrid.ItemsSource = models;
        }

        private RecordResultCategory GetSelectedCategory()
        {
            if (CategoryTargetOnlyRadio.IsChecked == true) return RecordResultCategory.TargetOnly;
            if (CategoryMatchingRadio.IsChecked == true) return RecordResultCategory.Matching;
            if (CategoryDuplicatesRadio.IsChecked == true) return RecordResultCategory.SourceDuplicates;
            return RecordResultCategory.SourceOnly;
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null) return;

            RecordResultCategory category = GetSelectedCategory();
            var matrix = _lastResult.ExportToMatrix(category);

            string sheetName = $"Reconcile_{category}";
            bool created = _context.CreateWorksheetWithData(sheetName, matrix);

            if (created)
            {
                _context.ShowMessage($"Exported {matrix.GetLength(0) - 1} records to worksheet '{sheetName}'.", "Export Completed");
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
