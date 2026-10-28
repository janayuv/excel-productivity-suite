using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FlipRangeDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;

        public FlipRangeDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (_matrix == null || PreviewTable == null) return;

            var previews = new List<PreviewItem>();
            try
            {
                bool isHorizontal = HorizontalRadio?.IsChecked == true;
                var flipped = isHorizontal
                    ? RangeFlipper.FlipHorizontal(_matrix)
                    : RangeFlipper.FlipVertical(_matrix);

                int rStart = _matrix.GetLowerBound(0);
                int cStart = _matrix.GetLowerBound(1);

                int limit = Math.Min(10, isHorizontal ? _matrix.GetLength(1) : _matrix.GetLength(0));

                for (int i = 0; i < limit; i++)
                {
                    string orig = isHorizontal
                        ? _matrix[rStart, cStart + i]?.ToString() ?? ""
                        : _matrix[rStart + i, cStart]?.ToString() ?? "";

                    string result = isHorizontal
                        ? flipped[0, i]?.ToString() ?? ""
                        : flipped[i, 0]?.ToString() ?? "";

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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before flipping ranges.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range to flip.", "Multi-Area Selection", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Flip Range", isError: true);
                return;
            }

            bool isHorizontal = HorizontalRadio?.IsChecked == true;
            var flipped = isHorizontal
                ? RangeFlipper.FlipHorizontal(_matrix)
                : RangeFlipper.FlipVertical(_matrix);

            _context.SetSelectedValues(flipped);
            DialogResult = true;
            Close();
            _context.ShowMessage($"Successfully flipped range {(isHorizontal ? "horizontally" : "vertically")}.", "Flip Range");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
