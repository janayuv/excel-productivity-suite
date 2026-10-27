using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class RemoveCharactersDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public RemoveCharactersDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private CharacterRemovalOptions GetOptions()
        {
            return new CharacterRemovalOptions
            {
                RemoveNumeric = RemoveNumericCheck.IsChecked == true,
                RemoveAlphabetic = RemoveAlphaCheck.IsChecked == true,
                RemoveNonAlphanumeric = RemoveSymbolsCheck.IsChecked == true,
                RemoveNonPrintable = RemoveControlCheck.IsChecked == true,
                RemoveCustomCharacters = RemoveCustomCheck.IsChecked == true,
                CustomCharacters = CustomCharsTextBox.Text
            };
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            if (CustomCharsTextBox != null)
            {
                CustomCharsTextBox.IsEnabled = RemoveCustomCheck.IsChecked == true;
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
                string result = CharacterRemover.RemoveCharacters(original, options);

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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before modifying characters.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells to modify.", "Remove Characters", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Remove Characters", isError: true);
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
                        string cleaned = CharacterRemover.RemoveCharacters(str, options);
                        updated[r, c] = cleaned;
                        affected++;
                    }
                }
            }

            _context.SetSelectedValues(updated);
            _context.ShowMessage($"Removed characters across {affected} cells.", "Remove Characters");
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
