using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ExtractTextDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public ExtractTextDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeInfoText.Text = $"Range: {_context.SelectionAddress ?? "No Selection"}";
            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private TextExtractionOptions GetOptions()
        {
            var options = new TextExtractionOptions
            {
                PrimaryDelimiter = PrimaryDelimTextBox.Text,
                SecondaryDelimiter = SecondaryDelimTextBox.Text,
                CaseSensitive = CaseSensitiveCheckBox.IsChecked == true,
                ReturnEmptyIfNotFound = true
            };

            if (ModeBefore.IsChecked == true) options.Mode = ExtractionMode.BeforeDelimiter;
            else if (ModeAfter.IsChecked == true) options.Mode = ExtractionMode.AfterDelimiter;
            else if (ModeBetween.IsChecked == true) options.Mode = ExtractionMode.BetweenDelimiters;

            return options;
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            if (SecondaryDelimTextBox != null)
            {
                bool isBetween = ModeBetween.IsChecked == true;
                SecondaryDelimTextBox.IsEnabled = isBetween;
                SecondaryDelimLabel.Foreground = isBetween ? (System.Windows.Media.Brush)FindResource("TextPrimaryBrush") : (System.Windows.Media.Brush)FindResource("TextSecondaryBrush");
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
                string result = TextExtractor.Extract(original, options);
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
            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells are currently selected in Excel.", "Extract Text", isError: true);
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
                        string extracted = TextExtractor.Extract(str, options);
                        updated[r, c] = extracted;
                        affected++;
                    }
                }
            }

            _context.SetSelectedValues(updated);
            _context.ShowMessage($"Extracted text successfully across {affected} cells.", "Extract Text");
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
