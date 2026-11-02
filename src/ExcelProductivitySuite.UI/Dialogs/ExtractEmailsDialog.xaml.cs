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
    public partial class ExtractEmailsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public ExtractEmailsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeInfoText.Text = $"Range: {_context.SelectionAddress ?? "No Active Selection"}";
            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private EmailExtractionOptions GetOptions()
        {
            var options = new EmailExtractionOptions();

            options.Mode = AllEmailsRadio?.IsChecked == true 
                ? EmailExtractionMode.AllJoined 
                : EmailExtractionMode.FirstOnly;

            options.Deduplicate = DeduplicateCheck?.IsChecked == true;
            options.Delimiter = DelimiterBox?.Text ?? "; ";

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
                string result = EmailExtractor.ExtractFromText(original, options);
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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before extracting emails.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells for email extraction.", "Extract Emails", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null || matrix.Length == 0)
            {
                _context.ShowMessage("No cells or empty range selected.", "Extract Emails", isError: true);
                return;
            }

            var options = GetOptions();
            var resultMatrix = EmailExtractor.ExtractFromMatrix(matrix, options);
            _context.SetSelectedValues(resultMatrix);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
