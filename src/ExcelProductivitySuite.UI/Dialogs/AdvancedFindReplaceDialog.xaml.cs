using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class AdvancedFindReplaceDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private List<FindMatchResult> _currentMatches = new();

        public AdvancedFindReplaceDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();
        }

        private FindReplaceOptions GetOptions()
        {
            return new FindReplaceOptions
            {
                SearchText = FindBox?.Text ?? "",
                ReplaceText = ReplaceBox?.Text ?? "",
                MatchCase = MatchCaseCheck?.IsChecked == true,
                MatchEntireCell = EntireCellCheck?.IsChecked == true,
                UseRegex = RegexCheck?.IsChecked == true
            };
        }

        private void FindAllButton_Click(object sender, RoutedEventArgs e)
        {
            var options = GetOptions();
            if (string.IsNullOrEmpty(options.SearchText))
            {
                _context.ShowMessage("Please enter text to search for.", "Find & Replace", isError: true);
                return;
            }

            _currentMatches.Clear();

            var sheetNames = AllSheetsRadio.IsChecked == true
                ? _context.GetWorksheetNames()
                : new List<string> { _context.ActiveWorksheetName ?? "Sheet1" };

            foreach (var sheet in sheetNames)
            {
                var matrix = _context.GetWorkbookRangeValues("", sheet);
                if (matrix != null)
                {
                    int rStart = matrix.GetLowerBound(0);
                    int cStart = matrix.GetLowerBound(1);
                    var matches = MultiSheetFindReplace.FindMatches(sheet, rStart, cStart, matrix, options);
                    _currentMatches.AddRange(matches);
                }
            }

            ResultsGrid.ItemsSource = null;
            ResultsGrid.ItemsSource = _currentMatches;
            ResultsStatusText.Text = $"Found {_currentMatches.Count} matching cell(s) across {sheetNames.Count} sheet(s).";
        }

        private void ReplaceAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("Worksheet protection is active. Please unprotect sheets before replacing.", "Sheet Protected", isError: true);
                return;
            }

            var options = GetOptions();
            if (string.IsNullOrEmpty(options.SearchText))
            {
                _context.ShowMessage("Please enter text to find and replace.", "Find & Replace", isError: true);
                return;
            }

            var sheetNames = AllSheetsRadio.IsChecked == true
                ? _context.GetWorksheetNames()
                : new List<string> { _context.ActiveWorksheetName ?? "Sheet1" };

            int totalReplacements = 0;

            foreach (var sheet in sheetNames)
            {
                var matrix = _context.GetWorkbookRangeValues("", sheet);
                if (matrix != null)
                {
                    var updated = MultiSheetFindReplace.ReplaceInMatrix(matrix, options, out int count);
                    if (count > 0)
                    {
                        int rStart = matrix.GetLowerBound(0);
                        int cStart = matrix.GetLowerBound(1);
                        _context.SetRangeValues(rStart, cStart, updated);
                        totalReplacements += count;
                    }
                }
            }

            _context.ShowMessage($"Replaced {totalReplacements} cell value(s) across {sheetNames.Count} sheet(s).", "Find & Replace");
            FindAllButton_Click(sender, e); // Refresh grid
        }

        private void SelectMatchesButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentMatches.Count == 0)
            {
                _context.ShowMessage("No search matches found to select.", "Find & Replace");
                return;
            }

            string activeSheet = _context.ActiveWorksheetName ?? "";
            var activeSheetCoords = _currentMatches
                .Where(m => string.Equals(m.SheetName, activeSheet, StringComparison.OrdinalIgnoreCase))
                .Select(m => (m.Row, m.Column))
                .ToList();

            if (activeSheetCoords.Count > 0)
            {
                _context.SelectCells(activeSheetCoords);
                _context.ShowMessage($"Selected {activeSheetCoords.Count} matching cell(s) on '{activeSheet}'.", "Find & Replace");
            }
            else
            {
                _context.ShowMessage("Matches found are on other sheets. Please switch to the target sheet to select them.", "Find & Replace");
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
