using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ChangeSignDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public ChangeSignDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private SignChangeOptions GetOptions()
        {
            var options = new SignChangeOptions();
            if (InvertRadio?.IsChecked == true)
            {
                options.Mode = SignChangeMode.InvertSign;
            }
            else if (MakePositiveRadio?.IsChecked == true)
            {
                options.Mode = SignChangeMode.MakeAllPositive;
            }
            else if (MakeNegativeRadio?.IsChecked == true)
            {
                options.Mode = SignChangeMode.MakeAllNegative;
            }
            else if (ReverseNegativeRadio?.IsChecked == true)
            {
                options.Mode = SignChangeMode.ReverseNegativeOnly;
            }
            else if (ReversePositiveRadio?.IsChecked == true)
            {
                options.Mode = SignChangeMode.ReversePositiveOnly;
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
                object? result = SignChanger.ChangeSign(original, options);

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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before changing signs.", "Sheet Protected", isError: true);
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
                _context.ShowMessage("No cells selected.", "Change Sign", isError: true);
                return;
            }

            var options = GetOptions();
            var updated = SignChanger.ChangeSignRange(matrix, options, out int affected);

            _context.SetSelectedValues(updated);
            DialogResult = true;
            Close();
            _context.ShowMessage($"Updated signs in {affected} numeric cells.", "Change Sign");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
