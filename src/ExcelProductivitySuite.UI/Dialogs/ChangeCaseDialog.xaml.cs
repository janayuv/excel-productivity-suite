using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ChangeCaseDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public ChangeCaseDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private TextCaseMode GetMode()
        {
            if (UpperRadio.IsChecked == true) return TextCaseMode.UpperCase;
            if (LowerRadio.IsChecked == true) return TextCaseMode.LowerCase;
            if (ProperRadio.IsChecked == true) return TextCaseMode.ProperCase;
            if (SentenceRadio.IsChecked == true) return TextCaseMode.SentenceCase;
            if (ToggleRadio.IsChecked == true) return TextCaseMode.ToggleCase;
            return TextCaseMode.UpperCase;
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTable == null) return;

            var mode = GetMode();
            var previews = new List<PreviewItem>();

            for (int i = 0; i < _sampleOriginals.Count; i++)
            {
                string original = _sampleOriginals[i];
                string result = CaseConverter.ConvertCase(original, mode);

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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before changing case.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells to modify.", "Change Case", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Change Case", isError: true);
                return;
            }

            var mode = GetMode();
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
                        string transformed = CaseConverter.ConvertCase(str, mode);
                        updated[r, c] = transformed;
                        affected++;
                    }
                }
            }

            _context.SetSelectedValues(updated);
            _context.ShowMessage($"Updated text case across {affected} cells.", "Change Case");
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
