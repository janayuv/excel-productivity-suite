using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class SwapRangesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public SwapRangesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            InitializeAddresses();
        }

        private void InitializeAddresses()
        {
            var areas = _context.GetSelectedAreaAddresses();
            if (areas != null && areas.Count >= 2)
            {
                Range1Box.Text = areas[0];
                Range2Box.Text = areas[1];
            }
            else
            {
                Range1Box.Text = _context.SelectionAddress ?? "$A$1:$B$5";
                Range2Box.Text = "$D$1:$E$5";
            }

            ValidateAndShowDimensions();
        }

        private void RangeBox_TextChanged(object sender, RoutedEventArgs e)
        {
            ValidateAndShowDimensions();
        }

        private void ValidateAndShowDimensions()
        {
            if (StatusInfoText == null) return;

            string addr1 = Range1Box?.Text?.Trim() ?? "";
            string addr2 = Range2Box?.Text?.Trim() ?? "";

            if (string.IsNullOrEmpty(addr1) || string.IsNullOrEmpty(addr2))
            {
                StatusInfoText.Text = "Please specify both Range 1 and Range 2 addresses.";
                return;
            }

            var r1 = _context.GetRangeValuesByAddress(addr1);
            var r2 = _context.GetRangeValuesByAddress(addr2);

            if (r1 == null || r2 == null)
            {
                StatusInfoText.Text = "One or both range addresses could not be resolved in the active workbook.";
                return;
            }

            int rows1 = r1.GetUpperBound(0) - r1.GetLowerBound(0) + 1;
            int cols1 = r1.GetUpperBound(1) - r1.GetLowerBound(1) + 1;

            int rows2 = r2.GetUpperBound(0) - r2.GetLowerBound(0) + 1;
            int cols2 = r2.GetUpperBound(1) - r2.GetLowerBound(1) + 1;

            bool matches = (rows1 == rows2 && cols1 == cols2);
            if (matches)
            {
                StatusInfoText.Text = $"Valid ranges ready to swap: {rows1} rows × {cols1} cols each.";
            }
            else
            {
                StatusInfoText.Text = $"Range 1 is {rows1}×{cols1}; Range 2 is {rows2}×{cols2}. " +
                    (RequireEqualDimensionsCheck.IsChecked == true ? "Warning: Dimensions do not match!" : "Will swap mismatched dimensions.");
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before swapping ranges.", "Sheet Protected", isError: true);
                return;
            }

            string addr1 = Range1Box.Text?.Trim() ?? "";
            string addr2 = Range2Box.Text?.Trim() ?? "";

            if (string.IsNullOrEmpty(addr1) || string.IsNullOrEmpty(addr2))
            {
                _context.ShowMessage("Please specify both Range 1 and Range 2 addresses.", "Swap Ranges", isError: true);
                return;
            }

            var r1 = _context.GetRangeValuesByAddress(addr1);
            var r2 = _context.GetRangeValuesByAddress(addr2);

            if (r1 == null || r2 == null)
            {
                _context.ShowMessage("Failed to read values from one or both range addresses.", "Swap Ranges", isError: true);
                return;
            }

            bool requireEqual = RequireEqualDimensionsCheck.IsChecked == true;
            try
            {
                var (newA, newB) = RangeSwapper.Swap(r1, r2, requireEqual);

                _context.SetRangeValuesByAddress(addr1, newA);
                _context.SetRangeValuesByAddress(addr2, newB);

                DialogResult = true;
                Close();
                _context.ShowMessage($"Successfully swapped data between '{addr1}' and '{addr2}'.", "Swap Ranges");
            }
            catch (Exception ex)
            {
                _context.ShowMessage(ex.Message, "Swap Error", isError: true);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
