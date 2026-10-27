using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class AddTextDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public AddTextDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private void ModeChanged(object sender, RoutedEventArgs e)
        {
            if (AffixPanel == null || PositionPanel == null) return;

            bool isAffix = ModeAffixesRadio.IsChecked == true;
            AffixPanel.Visibility = isAffix ? Visibility.Visible : Visibility.Collapsed;
            PositionPanel.Visibility = isAffix ? Visibility.Collapsed : Visibility.Visible;

            UpdatePreview();
        }

        private AffixOptions GetOptions()
        {
            return new AffixOptions
            {
                Prefix = PrefixTextBox.Text,
                Suffix = SuffixTextBox.Text,
                SkipBlankCells = SkipBlankCheckBox.IsChecked == true,
                SkipIfAlreadyContains = SkipIfContainsCheckBox.IsChecked == true
            };
        }

        private PositionInsertOptions GetPositionOptions()
        {
            int.TryParse(PositionIndexTextBox.Text, out int pos);
            if (pos < 1) pos = 1;

            return new PositionInsertOptions
            {
                TextToInsert = InsertTextBox.Text,
                PositionIndex = pos,
                FromEnd = FromEndCheckBox.IsChecked == true
            };
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTable == null) return;

            bool isAffix = ModeAffixesRadio?.IsChecked == true;
            var affixOpts = isAffix ? GetOptions() : null;
            var posOpts = !isAffix ? GetPositionOptions() : null;
            var previews = new List<PreviewItem>();

            for (int i = 0; i < _sampleOriginals.Count; i++)
            {
                string original = _sampleOriginals[i];
                string result = isAffix 
                    ? AffixAppender.ApplyAffixes(original, affixOpts!)
                    : PositionInserter.Insert(original, posOpts!);

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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before adding text.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells to modify.", "Add Text", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Add Text", isError: true);
                return;
            }

            bool isAffix = ModeAffixesRadio.IsChecked == true;
            var affixOpts = isAffix ? GetOptions() : null;
            var posOpts = !isAffix ? GetPositionOptions() : null;

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
                    string str = val?.ToString() ?? "";

                    if (isAffix)
                    {
                        if (!string.IsNullOrEmpty(str) || !affixOpts!.SkipBlankCells)
                        {
                            string result = AffixAppender.ApplyAffixes(str, affixOpts!);
                            updated[r, c] = result;
                            affected++;
                        }
                    }
                    else
                    {
                        string result = PositionInserter.Insert(str, posOpts!);
                        updated[r, c] = result;
                        affected++;
                    }
                }
            }

            _context.SetSelectedValues(updated);
            _context.ShowMessage($"Updated text across {affected} cells.", "Add Text");
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
