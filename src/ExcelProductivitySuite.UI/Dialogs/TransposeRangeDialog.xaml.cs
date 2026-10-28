using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class TransposeRangeDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;

        public TransposeRangeDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdateDimensions();
            UpdatePreview();
        }

        private void UpdateDimensions()
        {
            if (_matrix == null)
            {
                DimensionInfoText.Text = "No range selected.";
                return;
            }

            int rows = _matrix.GetUpperBound(0) - _matrix.GetLowerBound(0) + 1;
            int cols = _matrix.GetUpperBound(1) - _matrix.GetLowerBound(1) + 1;

            DimensionInfoText.Text = $"Source Selection: {rows} rows × {cols} cols  ➔  Transposed: {cols} rows × {rows} cols";
        }

        private void UpdatePreview()
        {
            if (_matrix == null || PreviewTable == null) return;

            var previews = new List<PreviewItem>();
            try
            {
                var transposed = RangeTransposer.Transpose(_matrix);
                int rStart = _matrix.GetLowerBound(0);
                int cStart = _matrix.GetLowerBound(1);

                int maxPreview = Math.Min(10, Math.Min(_matrix.GetLength(0), _matrix.GetLength(1)));
                if (maxPreview <= 0) maxPreview = Math.Min(10, Math.Max(_matrix.GetLength(0), _matrix.GetLength(1)));

                for (int i = 0; i < maxPreview; i++)
                {
                    string orig = "";
                    string trans = "";

                    if (i < _matrix.GetLength(0) && _matrix.GetLength(1) > 0)
                        orig = _matrix[rStart + i, cStart]?.ToString() ?? "";

                    if (i < transposed.GetLength(0) && transposed.GetLength(1) > 0)
                        trans = transposed[i, 0]?.ToString() ?? "";

                    previews.Add(new PreviewItem
                    {
                        RowIndex = i + 1,
                        Original = orig,
                        Result = trans
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
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before transposing.", "Sheet Protected", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Transpose Range", isError: true);
                return;
            }

            var transposed = RangeTransposer.Transpose(_matrix);

            if (NewSheetRadio.IsChecked == true)
            {
                string sheetName = SheetNameBox.Text?.Trim() ?? "Transposed";
                bool created = _context.CreateWorksheetWithData(sheetName, transposed);
                if (!created)
                {
                    _context.ShowMessage($"Could not create sheet '{sheetName}'. A sheet with this name may already exist.", "Transpose Range", isError: true);
                    return;
                }
            }
            else if (TargetAddressRadio.IsChecked == true)
            {
                string targetAddress = TargetAddressBox.Text?.Trim() ?? "$A$1";
                _context.SetRangeValuesByAddress(targetAddress, transposed);
            }
            else // InPlace
            {
                int startRow = _context.SelectionStartRow;
                int startCol = _context.SelectionStartColumn;
                _context.ClearSelection();
                _context.SetRangeValues(startRow, startCol, transposed);
            }

            DialogResult = true;
            Close();
            _context.ShowMessage($"Successfully transposed {transposed.GetLength(0)} × {transposed.GetLength(1)} range.", "Transpose Range");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
