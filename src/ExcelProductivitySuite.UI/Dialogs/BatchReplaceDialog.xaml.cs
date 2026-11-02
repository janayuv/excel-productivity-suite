using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class BatchReplaceDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string>? _sampleOriginals;
        private readonly ObservableCollection<BatchReplaceRule> _rules = new();

        public BatchReplaceDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeInfoText.Text = $"Range: {_context.SelectionAddress ?? "No Active Selection"}";
            _sampleOriginals = _context.GetSampleOriginalValues(10);

            RulesGrid.ItemsSource = _rules;
            _rules.CollectionChanged += (s, e) => UpdatePreview();

            // Populate starter sample rule
            _rules.Add(new BatchReplaceRule("St.", "Street"));
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTable == null || _sampleOriginals == null) return;

            var previews = new List<PreviewItem>();

            for (int i = 0; i < _sampleOriginals.Count; i++)
            {
                string original = _sampleOriginals[i];
                string result = BatchReplacer.ReplaceInText(original, _rules, out _);
                previews.Add(new PreviewItem
                {
                    RowIndex = i + 1,
                    Original = original,
                    Result = result
                });
            }

            PreviewTable.ItemsSource = previews;
        }

        private void AddRuleButton_Click(object sender, RoutedEventArgs e)
        {
            string find = FindInputBox.Text ?? "";
            if (string.IsNullOrEmpty(find))
            {
                _context.ShowMessage("Please enter a text string to find.", "Batch Replace", isError: true);
                return;
            }

            string replaceWith = ReplaceInputBox.Text ?? "";
            bool matchCase = MatchCaseCheck.IsChecked == true;
            bool entireCell = EntireCellCheck.IsChecked == true;
            bool useRegex = UseRegexCheck.IsChecked == true;

            _rules.Add(new BatchReplaceRule(find, replaceWith, matchCase, entireCell, useRegex));
            FindInputBox.Clear();
            ReplaceInputBox.Clear();
            UpdatePreview();
        }

        private void RemoveRuleButton_Click(object sender, RoutedEventArgs e)
        {
            if (RulesGrid.SelectedItem is BatchReplaceRule rule)
            {
                _rules.Remove(rule);
                UpdatePreview();
            }
        }

        private void ClearAllRulesButton_Click(object sender, RoutedEventArgs e)
        {
            _rules.Clear();
            UpdatePreview();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before replacing text.", "Sheet Protected", isError: true);
                return;
            }

            if (_rules.Count == 0)
            {
                _context.ShowMessage("Please define at least one replacement rule.", "Batch Replace", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null || matrix.Length == 0)
            {
                _context.ShowMessage("No cells or empty range selected.", "Batch Replace", isError: true);
                return;
            }

            var resultMatrix = BatchReplacer.ReplaceInMatrix(matrix, _rules, out int totalReplacements);
            _context.SetSelectedValues(resultMatrix);
            _context.ShowMessage($"Batch replace complete. Replaced {totalReplacements} occurrences across the selection.", "Batch Replace", isError: false);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
