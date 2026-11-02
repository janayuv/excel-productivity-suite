using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ExtractNumbersDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public ExtractNumbersDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeInfoText.Text = $"Range: {_context.SelectionAddress ?? "No Active Selection"}";
            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private NumberExtractionOptions GetOptions()
        {
            var options = new NumberExtractionOptions();

            if (FirstOnlyRadio?.IsChecked == true) options.Mode = NumberExtractionMode.FirstOnly;
            else if (SumRadio?.IsChecked == true) options.Mode = NumberExtractionMode.Sum;
            else options.Mode = NumberExtractionMode.AllJoined;

            options.AllowDecimals = AllowDecimalsCheck?.IsChecked == true;
            options.AllowNegative = AllowNegativeCheck?.IsChecked == true;
            options.Delimiter = DelimiterBox?.Text ?? ", ";

            return options;
        }

        private void UpdatePreview()
        {
            if (PreviewTable == null || _sampleOriginals == null) return;

            var options = GetOptions();
            var previews = new List<PreviewItem>();

            for (int i = 0; i < _sampleOriginals.Count; i++)
            {
                string original = _sampleOriginals[i];
                string result = NumberExtractor.ExtractFromText(original, options);
                previews.Add(new PreviewItem
                {
                    RowIndex = i + 1,
                    Original = original,
                    Result = result
                });
            }

            PreviewTable.ItemsSource = previews;
        }

        private void Option_Changed(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void Delimiter_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before extracting numbers.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells for number extraction.", "Extract Numbers", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null || matrix.Length == 0)
            {
                _context.ShowMessage("No cells or empty range selected.", "Extract Numbers", isError: true);
                return;
            }

            var options = GetOptions();
            var resultMatrix = NumberExtractor.ExtractFromMatrix(matrix, options);
            _context.SetSelectedValues(resultMatrix);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
