using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class TransformRangeDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;
        private readonly bool _isInitializing = true;

        public TransformRangeDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            _isInitializing = false;

            UpdateUiState();
            UpdatePreview();
        }

        private void UpdateUiState()
        {
            if (rbColToRange == null || panelListToRange == null || panelRangeToList == null) return;

            bool isListToRange = rbColToRange.IsChecked == true || rbRowToRange.IsChecked == true;
            panelListToRange.Visibility = isListToRange ? Visibility.Visible : Visibility.Collapsed;
            panelRangeToList.Visibility = isListToRange ? Visibility.Collapsed : Visibility.Visible;
        }

        private TransformRangeOptions BuildOptions()
        {
            var options = new TransformRangeOptions();

            if (rbColToRange.IsChecked == true)
                options.Type = TransformType.SingleColumnToRange;
            else if (rbRowToRange.IsChecked == true)
                options.Type = TransformType.SingleRowToRange;
            else if (rbRangeToCol.IsChecked == true)
                options.Type = TransformType.RangeToSingleColumn;
            else
                options.Type = TransformType.RangeToSingleRow;

            options.RecordSplit = (rbBlankDelimited.IsChecked == true)
                ? TransformRecordSplit.BlankCellDelimited
                : TransformRecordSplit.FixedValue;

            if (int.TryParse(txtColumnsPerRecord?.Text, out int cols) && cols >= 1)
                options.ColumnsPerRecord = cols;
            else
                options.ColumnsPerRecord = 5;

            options.Order = (rbOrderByCol.IsChecked == true)
                ? TransformOrder.ByColumn
                : TransformOrder.ByRow;

            options.SkipBlankCells = (chkSkipBlanks?.IsChecked == true);

            if (rbInPlace?.IsChecked == true)
                options.DestinationMode = TransformDestinationMode.InPlace;
            else if (rbNewSheet?.IsChecked == true)
                options.DestinationMode = TransformDestinationMode.NewSheet;
            else
                options.DestinationMode = TransformDestinationMode.TargetCell;

            options.TargetAddress = string.IsNullOrWhiteSpace(txtTargetCell?.Text) ? "$E$1" : (txtTargetCell?.Text?.Trim() ?? "$E$1");
            options.NewSheetName = string.IsNullOrWhiteSpace(txtSheetName?.Text) ? "Transformed" : (txtSheetName?.Text?.Trim() ?? "Transformed");

            return options;
        }

        private void UpdatePreview()
        {
            if (_isInitializing || _matrix == null || PreviewTable == null || DimensionInfoText == null) return;

            int srcRows = _matrix.GetUpperBound(0) - _matrix.GetLowerBound(0) + 1;
            int srcCols = _matrix.GetUpperBound(1) - _matrix.GetLowerBound(1) + 1;

            var previews = new List<PreviewItem>();
            try
            {
                var options = BuildOptions();
                var transformed = RangeTransformer.Transform(_matrix, options);

                int outRows = transformed.GetLength(0);
                int outCols = transformed.GetLength(1);

                DimensionInfoText.Text = $"Source Selection: {srcRows} rows × {srcCols} cols  ➔  Transformed: {outRows} rows × {outCols} cols";

                int maxPreview = Math.Min(10, Math.Max(srcRows, outRows));
                int rStart = _matrix.GetLowerBound(0);
                int cStart = _matrix.GetLowerBound(1);

                for (int i = 0; i < maxPreview; i++)
                {
                    string orig = "";
                    string trans = "";

                    if (i < srcRows && srcCols > 0)
                    {
                        orig = _matrix[rStart + i, cStart]?.ToString() ?? "";
                        if (srcCols > 1)
                        {
                            orig += $" (row {i + 1})";
                        }
                    }

                    if (i < outRows && outCols > 0)
                    {
                        if (outCols == 1)
                        {
                            trans = transformed[i, 0]?.ToString() ?? "";
                        }
                        else
                        {
                            var parts = new List<string>();
                            for (int c = 0; c < Math.Min(4, outCols); c++)
                            {
                                parts.Add(transformed[i, c]?.ToString() ?? "");
                            }
                            if (outCols > 4) parts.Add("...");
                            trans = string.Join(" | ", parts);
                        }
                    }

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
                DimensionInfoText.Text = "Please verify transformation settings.";
                previews.Add(new PreviewItem { RowIndex = 1, Original = "Status", Result = ex.Message });
            }

            PreviewTable.ItemsSource = previews;
        }

        private void OnTypeChanged(object sender, RoutedEventArgs e)
        {
            UpdateUiState();
            UpdatePreview();
        }

        private void OnOptionChanged(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void OnOptionChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before transforming.", "Sheet Protected", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Transform Range", isError: true);
                return;
            }

            var options = BuildOptions();
            var result = _context.TransformRange(options);

            if (result.Success)
            {
                DialogResult = true;
                Close();
                _context.ShowMessage(result.Message, "Transform Range");
            }
            else
            {
                _context.ShowMessage(result.Message, "Transform Range", isError: true);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
