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
    public partial class ConvertDatesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public ConvertDatesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private DateTextConversionOptions GetOptions()
        {
            var options = new DateTextConversionOptions();

            if (FormatComboBox?.SelectedItem is ComboBoxItem fItem &&
                Enum.TryParse<DateOutputFormat>(fItem.Tag?.ToString(), out var format))
            {
                options.OutputFormat = format;
            }

            if (CultureComboBox?.SelectedItem is ComboBoxItem cItem &&
                Enum.TryParse<DateCulturePreference>(cItem.Tag?.ToString(), out var cult))
            {
                options.CulturePreference = cult;
            }

            options.SkipUnparseable = SkipUnparseableCheck?.IsChecked == true;
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
                object? converted = DateTextConverter.ConvertToDate(original, options);

                previews.Add(new PreviewItem
                {
                    RowIndex = i + 1,
                    Original = original,
                    Result = converted?.ToString() ?? ""
                });
            }

            PreviewTable.ItemsSource = previews;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before converting dates.", "Sheet Protected", isError: true);
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
                _context.ShowMessage("No cells selected.", "Convert Dates", isError: true);
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
                        object? converted = DateTextConverter.ConvertToDate(val, options);
                        if (converted != null && !Equals(converted, val))
                        {
                            updated[r, c] = converted;
                            affected++;
                        }
                    }
                }
            }

            _context.SetSelectedValues(updated);
            DialogResult = true;
            Close();
            _context.ShowMessage($"Converted {affected} text cells to formatted dates.", "Convert Dates");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
