using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class RandomSortDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;

        public RandomSortDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdatePreview();
        }

        private RandomSortOptions GetOptions()
        {
            var options = new RandomSortOptions
            {
                HasHeaderRow = HeaderRowCheck?.IsChecked == true
            };

            if (EntireRowsRadio?.IsChecked == true)
            {
                options.Mode = RandomSortMode.EntireRows;
            }
            else if (CellsWithinColsRadio?.IsChecked == true)
            {
                options.Mode = RandomSortMode.CellsWithinColumns;
            }
            else if (EntireColsRadio?.IsChecked == true)
            {
                options.Mode = RandomSortMode.EntireColumns;
            }
            else if (CellsWithinRowsRadio?.IsChecked == true)
            {
                options.Mode = RandomSortMode.CellsWithinRows;
            }

            return options;
        }

        private void UpdatePreview()
        {
            if (_matrix == null || PreviewTable == null) return;

            var previews = new List<PreviewItem>();
            try
            {
                var options = GetOptions();
                // Use fixed seed for preview determinism
                var shuffled = RangeRandomizer.Randomize(_matrix, options, new Random(42));

                int rStart = _matrix.GetLowerBound(0);
                int cStart = _matrix.GetLowerBound(1);
                int limit = Math.Min(10, _matrix.GetLength(0));

                for (int i = 0; i < limit; i++)
                {
                    string orig = _matrix[rStart + i, cStart]?.ToString() ?? "";
                    string result = shuffled[rStart + i, cStart]?.ToString() ?? "";

                    previews.Add(new PreviewItem
                    {
                        RowIndex = i + 1,
                        Original = orig,
                        Result = result
                    });
                }
            }
            catch (Exception ex)
            {
                previews.Add(new PreviewItem { RowIndex = 1, Original = "Error", Result = ex.Message });
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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before sorting.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range.", "Multi-Area Selection", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Random Sort", isError: true);
                return;
            }

            var options = GetOptions();
            var shuffled = RangeRandomizer.Randomize(_matrix, options);

            _context.SetSelectedValues(shuffled);
            DialogResult = true;
            Close();
            _context.ShowMessage("Range randomized successfully.", "Random Sort");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
