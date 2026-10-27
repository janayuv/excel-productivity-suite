using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class WhitespaceDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public WhitespaceDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private WhitespaceOptions GetOptions()
        {
            var options = new WhitespaceOptions
            {
                ConvertNonBreakingSpaces = ConvertNBSPCheckBox.IsChecked == true
            };

            if (LeadingTrailingRadio.IsChecked == true) options.Mode = WhitespaceTrimMode.TrimLeadingAndTrailing;
            else if (RemoveAllRadio.IsChecked == true) options.Mode = WhitespaceTrimMode.RemoveAllSpaces;
            else options.Mode = WhitespaceTrimMode.CollapseExcessInternalSpaces;

            return options;
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
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
                string result = WhitespaceNormalizer.Normalize(original, options);

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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before trimming spaces.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells to modify.", "Trim Spaces", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Trim Spaces", isError: true);
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
                        string cleaned = WhitespaceNormalizer.Normalize(str, options);
                        updated[r, c] = cleaned;
                        affected++;
                    }
                }
            }

            _context.SetSelectedValues(updated);
            _context.ShowMessage($"Cleaned whitespace across {affected} cells.", "Trim Spaces");
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
