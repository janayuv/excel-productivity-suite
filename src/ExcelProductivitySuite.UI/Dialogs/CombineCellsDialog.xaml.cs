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
    public partial class CombineCellsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;

        public CombineCellsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdatePreview();
        }

        private CombineOptions GetOptions()
        {
            var options = new CombineOptions();

            if (AcrossColumnsRadio?.IsChecked == true)
            {
                options.Direction = CombineDirection.AcrossColumns;
            }
            else if (DownRowsRadio?.IsChecked == true)
            {
                options.Direction = CombineDirection.DownRows;
            }
            else if (EntireSelectionRadio?.IsChecked == true)
            {
                options.Direction = CombineDirection.EntireSelection;
            }

            options.SkipBlankCells = SkipBlankCheck?.IsChecked == true;
            options.TrimTokens = TrimTokensCheck?.IsChecked == true;

            if (SeparatorComboBox?.SelectedItem is ComboBoxItem item)
            {
                string tag = item.Tag?.ToString() ?? "";
                if (tag == "custom")
                {
                    options.Delimiter = CustomSeparatorTextBox?.Text ?? "";
                }
                else
                {
                    options.Delimiter = tag;
                }
            }

            return options;
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void SeparatorComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CustomSeparatorPanel != null && SeparatorComboBox?.SelectedItem is ComboBoxItem item)
            {
                bool isCustom = item.Tag?.ToString() == "custom";
                CustomSeparatorPanel.Visibility = isCustom ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTable == null || _matrix == null) return;

            var options = GetOptions();
            var previews = new List<PreviewItem>();

            try
            {
                var combinedMatrix = TextCombiner.CombineMatrix(_matrix, options);
                int rStart = combinedMatrix.GetLowerBound(0);
                int rEnd = combinedMatrix.GetUpperBound(0);
                int cStart = combinedMatrix.GetLowerBound(1);
                int cEnd = combinedMatrix.GetUpperBound(1);

                int maxSamples = 10;
                int count = 0;

                for (int r = rStart; r <= rEnd && count < maxSamples; r++)
                {
                    for (int c = cStart; c <= cEnd && count < maxSamples; c++)
                    {
                        string original = (r < _matrix.GetLength(0) && c < _matrix.GetLength(1))
                            ? (_matrix[r, c]?.ToString() ?? "")
                            : "";

                        string result = combinedMatrix[r, c]?.ToString() ?? "";

                        previews.Add(new PreviewItem
                        {
                            RowIndex = count + 1,
                            Original = original,
                            Result = result
                        });
                        count++;
                    }
                }
            }
            catch
            {
                // Fallback
            }

            PreviewTable.ItemsSource = previews;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before combining cells.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells to combine.", "Multi-Area Selection", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Combine Cells", isError: true);
                return;
            }

            var options = GetOptions();
            var combinedMatrix = TextCombiner.CombineMatrix(_matrix, options);

            bool outputAdjacent = OutputAdjacentCheck?.IsChecked == true;

            if (outputAdjacent)
            {
                int targetStartRow = _context.SelectionStartRow;
                int targetStartCol = _context.SelectionStartColumn + _context.SelectionColumnCount;

                int outRows = combinedMatrix.GetLength(0);
                int outCols = combinedMatrix.GetLength(1);

                if (_context.CheckRangeHasData(targetStartRow, targetStartCol, outRows, outCols))
                {
                    bool proceed = _context.Confirm(
                        "The destination range already contains data. Overwrite existing cells?",
                        "Confirm Overwrite");
                    if (!proceed) return;
                }

                _context.SetRangeValues(targetStartRow, targetStartCol, combinedMatrix);
            }
            else
            {
                // In-place assignment
                if (options.Direction == CombineDirection.AcrossColumns)
                {
                    // If rows match selection row count and writing into col 1
                    int targetStartRow = _context.SelectionStartRow;
                    int targetStartCol = _context.SelectionStartColumn;
                    _context.SetRangeValues(targetStartRow, targetStartCol, combinedMatrix);
                }
                else if (options.Direction == CombineDirection.DownRows)
                {
                    int targetStartRow = _context.SelectionStartRow;
                    int targetStartCol = _context.SelectionStartColumn;
                    _context.SetRangeValues(targetStartRow, targetStartCol, combinedMatrix);
                }
                else
                {
                    // Single cell: set top-left cell
                    int targetStartRow = _context.SelectionStartRow;
                    int targetStartCol = _context.SelectionStartColumn;
                    _context.SetRangeValues(targetStartRow, targetStartCol, combinedMatrix);
                }
            }

            DialogResult = true;
            Close();
            _context.ShowMessage("Successfully combined cells.", "Combine Cells");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
