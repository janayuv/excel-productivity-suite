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
    public partial class FindDuplicatesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;
        private DuplicateAnalysisResult? _currentResult;

        public FindDuplicatesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdateAnalysis();
        }

        private DuplicateAnalysisOptions GetOptions()
        {
            var options = new DuplicateAnalysisOptions
            {
                MatchCase = MatchCaseCheck?.IsChecked == true,
                HasHeaders = HasHeadersCheck?.IsChecked == true
            };

            if (ExcludeFirstRadio?.IsChecked == true)
            {
                options.MatchRule = DuplicateMatchRule.DuplicatesExcludingFirst;
            }
            else if (AllDuplicatesRadio?.IsChecked == true)
            {
                options.MatchRule = DuplicateMatchRule.AllDuplicates;
            }
            else if (UniqueValuesRadio?.IsChecked == true)
            {
                options.MatchRule = DuplicateMatchRule.UniqueValuesOnly;
            }

            if (EntireRowsRadio?.IsChecked == true)
            {
                options.Scope = DuplicateAnalysisScope.EntireRows;
            }
            else
            {
                options.Scope = DuplicateAnalysisScope.CellByCell;
            }

            return options;
        }

        private void UpdateAnalysis()
        {
            if (ResultSummaryText == null) return;
            if (_matrix == null)
            {
                ResultSummaryText.Text = "No range selected.";
                return;
            }

            try
            {
                var options = GetOptions();
                int startRow = _context.SelectionStartRow;
                int startCol = _context.SelectionStartColumn;

                _currentResult = DuplicateAnalyzer.Analyze(_matrix, startRow, startCol, options);

                int matchCount = _currentResult.MatchingCells.Count;
                string ruleName = options.MatchRule switch
                {
                    DuplicateMatchRule.DuplicatesExcludingFirst => "Duplicates (excl. 1st)",
                    DuplicateMatchRule.AllDuplicates => "All Duplicates",
                    DuplicateMatchRule.UniqueValuesOnly => "Unique Values",
                    _ => "Matches"
                };

                ResultSummaryText.Text = $"Found {matchCount} matching cells ({_currentResult.DuplicateGroupCount} repeat groups, {_currentResult.UniqueValueCount} unique values).";
                DetailSummaryText.Text = $"Rule: {ruleName} | Analyzed: {_currentResult.TotalCellsAnalyzed} cells in range.";
            }
            catch (Exception ex)
            {
                ResultSummaryText.Text = "Error: " + ex.Message;
            }
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdateAnalysis();
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentResult == null || _currentResult.MatchingCells.Count == 0)
            {
                _context.ShowMessage("No matching cells found to select.", "Find Duplicates");
                return;
            }

            var cellCoords = _currentResult.MatchingCells.Select(c => (c.Row, c.Column)).ToList();
            _context.SelectCells(cellCoords);

            DialogResult = true;
            Close();
            _context.ShowMessage($"Selected {cellCoords.Count} matching cell(s) in Excel.", "Find Duplicates");
        }

        private void HighlightButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before highlighting.", "Sheet Protected", isError: true);
                return;
            }

            if (_currentResult == null || _currentResult.MatchingCells.Count == 0)
            {
                _context.ShowMessage("No matching cells found to highlight.", "Find Duplicates");
                return;
            }

            int colorRgb = 65535; // Default yellow
            if (HighlightColorCombo.SelectedItem is ComboBoxItem item && int.TryParse(item.Tag?.ToString(), out int tagColor))
            {
                colorRgb = tagColor;
            }

            var cellCoords = _currentResult.MatchingCells.Select(c => (c.Row, c.Column)).ToList();
            _context.HighlightCells(cellCoords, colorRgb);

            DialogResult = true;
            Close();
            _context.ShowMessage($"Highlighted {cellCoords.Count} matching cell(s).", "Find Duplicates");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
