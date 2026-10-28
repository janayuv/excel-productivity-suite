using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class RoundValuesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public RoundValuesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private ValueRoundOptions GetOptions()
        {
            var options = new ValueRoundOptions();

            if (int.TryParse(DecimalsBox?.Text?.Trim(), out int dec))
            {
                options.DecimalPlaces = Math.Max(0, dec);
            }

            if (double.TryParse(MultipleBox?.Text?.Trim(), out double mult) && mult > 0)
            {
                options.Multiple = mult;
            }

            if (StandardRadio?.IsChecked == true)
            {
                options.Mode = RoundingMode.Standard;
            }
            else if (RoundUpRadio?.IsChecked == true)
            {
                options.Mode = RoundingMode.RoundUpCeiling;
            }
            else if (RoundDownRadio?.IsChecked == true)
            {
                options.Mode = RoundingMode.RoundDownFloor;
            }
            else if (TruncateRadio?.IsChecked == true)
            {
                options.Mode = RoundingMode.Truncate;
            }
            else if (NearestMultipleRadio?.IsChecked == true)
            {
                options.Mode = RoundingMode.NearestMultiple;
            }

            return options;
        }

        private void UpdatePreview()
        {
            if (PreviewTable == null) return;

            var options = GetOptions();
            var previews = new List<PreviewItem>();

            for (int i = 0; i < _sampleOriginals.Count; i++)
            {
                string original = _sampleOriginals[i];
                object? result = ValueRounder.Round(original, options);

                previews.Add(new PreviewItem
                {
                    RowIndex = i + 1,
                    Original = original,
                    Result = result?.ToString() ?? "(blank)"
                });
            }

            PreviewTable.ItemsSource = previews;
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before rounding.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range.", "Multi-Area Selection", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Round Values", isError: true);
                return;
            }

            var options = GetOptions();
            var updated = ValueRounder.RoundRange(matrix, options, out int affected);

            _context.SetSelectedValues(updated);
            DialogResult = true;
            Close();
            _context.ShowMessage($"Rounded {affected} numeric cells successfully.", "Round Values");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
