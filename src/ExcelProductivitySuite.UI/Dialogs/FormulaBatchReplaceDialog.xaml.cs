using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formulas;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FormulaBatchReplaceDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private FormulaBatchReplaceResult? _latestResult;
        private bool _isScopeSelection = true;

        public ObservableCollection<FormulaChangeItem> ChangeItems { get; } = new ObservableCollection<FormulaChangeItem>();

        public FormulaBatchReplaceDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            ChangesGrid.ItemsSource = ChangeItems;
        }

        private void PreviewBtn_Click(object sender, RoutedEventArgs e)
        {
            string findText = FindBox.Text;
            if (string.IsNullOrEmpty(findText))
            {
                _context.ShowMessage("Please specify a string or reference pattern to find.", "Validation Error", isError: true);
                return;
            }

            string replaceWith = ReplaceBox.Text ?? string.Empty;
            _isScopeSelection = ScopeSelectionRadio.IsChecked == true;

            var options = new FormulaReplaceOptions
            {
                FindText = findText,
                ReplaceWith = replaceWith,
                MatchCase = MatchCaseCheck.IsChecked == true,
                MatchWholeWord = WholeWordCheck.IsChecked == true,
                IgnoreStringLiterals = IgnoreQuotesCheck.IsChecked == true
            };

            object?[,]? formulas = _isScopeSelection
                ? _context.GetSelectedFormulas()
                : _context.GetRangeFormulas(1, 1, 1000, 100);

            if (formulas == null || formulas.Length == 0)
            {
                StatusText.Text = "No formulas found in the targeted range.";
                ChangeItems.Clear();
                ApplyBtn.IsEnabled = false;
                return;
            }

            int startR = _isScopeSelection ? _context.SelectionStartRow : 1;
            int startC = _isScopeSelection ? _context.SelectionStartColumn : 1;

            _latestResult = FormulaBatchReplacer.ProcessFormulas(formulas, options, startR, startC);

            ChangeItems.Clear();
            foreach (var chg in _latestResult.Changes)
            {
                ChangeItems.Add(chg);
            }

            if (_latestResult.FormulasModified > 0)
            {
                StatusText.Text = $"Found {_latestResult.TotalReplacements} replacement(s) across {_latestResult.FormulasModified} formula(s). Ready to apply.";
                ApplyBtn.IsEnabled = true;
            }
            else
            {
                StatusText.Text = $"No matches found in {_latestResult.FormulasFound} formula cells scanned.";
                ApplyBtn.IsEnabled = false;
            }
        }

        private void ApplyBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_latestResult == null || _latestResult.FormulasModified == 0) return;

            bool confirm = _context.Confirm(
                $"Apply {_latestResult.TotalReplacements} replacement(s) to {_latestResult.FormulasModified} formula(s)?",
                "Confirm Formula Replacement");

            if (!confirm) return;

            try
            {
                if (_isScopeSelection)
                {
                    _context.SetSelectedFormulas(_latestResult.UpdatedFormulas);
                }
                else
                {
                    _context.SetRangeFormulas(1, 1, _latestResult.UpdatedFormulas);
                }

                _context.ShowMessage(
                    $"Successfully updated {_latestResult.FormulasModified} formula(s) with {_latestResult.TotalReplacements} replacement(s).",
                    "Replacement Complete");

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Failed to update formulas: {ex.Message}", "Replacement Error", isError: true);
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
