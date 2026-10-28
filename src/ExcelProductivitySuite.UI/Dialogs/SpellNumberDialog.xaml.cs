using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class SpellNumberDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public SpellNumberDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private NumberSpellerOptions GetOptions()
        {
            var options = new NumberSpellerOptions();

            if (ModeComboBox?.SelectedItem is ComboBoxItem mItem &&
                Enum.TryParse<SpellerMode>(mItem.Tag?.ToString(), out var mode))
            {
                options.Mode = mode;
            }

            if (CurrencyComboBox?.SelectedItem is ComboBoxItem currItem &&
                Enum.TryParse<SpellerCurrency>(currItem.Tag?.ToString(), out var curr))
            {
                options.Currency = curr;
            }

            if (CaseComboBox?.SelectedItem is ComboBoxItem caseItem &&
                Enum.TryParse<SpellerCase>(caseItem.Tag?.ToString(), out var scase))
            {
                options.LetterCase = scase;
            }

            options.IncludeCentsIfZero = ZeroCentsCheck?.IsChecked == true;
            return options;
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            if (CurrencyComboBox != null && ModeComboBox?.SelectedItem is ComboBoxItem mItem)
            {
                CurrencyComboBox.IsEnabled = mItem.Tag?.ToString() == "Currency";
            }
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTable == null) return;

            var options = GetOptions();
            var previews = new List<PreviewItem>();

            for (int i = 0; i < _sampleOriginals.Count; i++)
            {
                string original = _sampleOriginals[i];
                string result = "";

                if (double.TryParse(original, NumberStyles.Any, CultureInfo.InvariantCulture, out double num) ||
                    double.TryParse(original, NumberStyles.Any, CultureInfo.CurrentCulture, out num))
                {
                    result = NumberSpeller.Spell(num, options);
                }
                else
                {
                    result = original; // Non-numeric left as is
                }

                previews.Add(new PreviewItem
                {
                    RowIndex = i + 1,
                    Original = original,
                    Result = result
                });
            }

            PreviewTable.ItemsSource = previews;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before converting numbers to words.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells to convert.", "Multi-Area Selection", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Spell Number", isError: true);
                return;
            }

            var options = GetOptions();
            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int affected = 0;
            var updated = (object?[,])matrix.Clone();

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    var val = matrix[r, c];
                    if (val != null)
                    {
                        string str = val.ToString() ?? "";
                        if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double num) ||
                            double.TryParse(str, NumberStyles.Any, CultureInfo.CurrentCulture, out num))
                        {
                            updated[r, c] = NumberSpeller.Spell(num, options);
                            affected++;
                        }
                    }
                }
            }

            bool outputAdjacent = OutputAdjacentCheck?.IsChecked == true;
            if (outputAdjacent)
            {
                int targetStartRow = _context.SelectionStartRow;
                int targetStartCol = _context.SelectionStartColumn + _context.SelectionColumnCount;
                int rows = updated.GetLength(0);
                int cols = updated.GetLength(1);

                if (_context.CheckRangeHasData(targetStartRow, targetStartCol, rows, cols))
                {
                    bool proceed = _context.Confirm("Destination range contains data. Overwrite?", "Confirm Overwrite");
                    if (!proceed) return;
                }

                _context.SetRangeValues(targetStartRow, targetStartCol, updated);
            }
            else
            {
                _context.SetSelectedValues(updated);
            }

            DialogResult = true;
            Close();
            _context.ShowMessage($"Converted {affected} numbers into written words.", "Spell Number");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
