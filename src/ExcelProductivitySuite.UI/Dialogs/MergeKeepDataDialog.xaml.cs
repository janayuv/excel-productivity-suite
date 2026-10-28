using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class MergeKeepDataDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;

        public MergeKeepDataDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdatePreview();
        }

        private SafeMergeOptions GetOptions()
        {
            var options = new SafeMergeOptions
            {
                SkipBlankCells = SkipBlankCellsCheck?.IsChecked == true
            };

            if (RowByRowRadio?.IsChecked == true)
            {
                options.Direction = SafeMergeDirection.RowByRow;
            }
            else if (ColByColRadio?.IsChecked == true)
            {
                options.Direction = SafeMergeDirection.ColumnByColumn;
            }
            else if (EntireSelectionRadio?.IsChecked == true)
            {
                options.Direction = SafeMergeDirection.EntireSelection;
            }

            if (SeparatorCombo?.SelectedItem is ComboBoxItem item)
            {
                string tag = item.Tag?.ToString() ?? " ";
                if (tag == "CUSTOM")
                {
                    options.Separator = CustomSeparatorBox?.Text ?? " ";
                }
                else
                {
                    options.Separator = tag;
                }
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
                var merged = SafeRangeMerger.PrepareMerge(_matrix, options, out int _);

                int rStart = _matrix.GetLowerBound(0);
                int cStart = _matrix.GetLowerBound(1);
                int limit = Math.Min(10, _matrix.GetLength(0));

                for (int i = 0; i < limit; i++)
                {
                    string orig = "";
                    for (int c = 0; c < _matrix.GetLength(1); c++)
                    {
                        var v = _matrix[rStart + i, cStart + c]?.ToString();
                        if (!string.IsNullOrEmpty(v)) orig += (orig.Length > 0 ? " | " : "") + v;
                    }

                    string result = merged[i, 0]?.ToString() ?? "";

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

        private void SeparatorCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CustomSeparatorBox != null)
            {
                bool isCustom = (SeparatorCombo.SelectedItem is ComboBoxItem item && item.Tag?.ToString() == "CUSTOM");
                CustomSeparatorBox.Visibility = isCustom ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdatePreview();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before merging.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range.", "Multi-Area Selection", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Merge Cells", isError: true);
                return;
            }

            var options = GetOptions();
            var prepared = SafeRangeMerger.PrepareMerge(_matrix, options, out int groups);

            _context.SetSelectedValues(prepared);
            DialogResult = true;
            Close();
            _context.ShowMessage($"Consolidated data across {groups} merged group(s) without losing information.", "Merge Cells (Keep Data)");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
