using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class CopyMultipleRangesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private IReadOnlyList<object?[,]> _areas = new List<object?[,]>();

        public class AreaViewModel
        {
            public string AreaNumber { get; }
            public string Address { get; }
            public string DimensionsText { get; }

            public AreaViewModel(int number, string address, int rows, int cols)
            {
                AreaNumber = $"Area {number}:";
                Address = address;
                DimensionsText = $"({rows} rows × {cols} cols)";
            }
        }

        public CopyMultipleRangesDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadAreas();
        }

        private void LoadAreas()
        {
            _areas = _context.GetMultiAreaSelectedValues();
            var addresses = _context.GetMultiAreaAddresses();

            var viewModels = new List<AreaViewModel>();
            for (int i = 0; i < _areas.Count; i++)
            {
                var a = _areas[i];
                int r = a.GetUpperBound(0) - a.GetLowerBound(0) + 1;
                int c = a.GetUpperBound(1) - a.GetLowerBound(1) + 1;
                string addr = i < addresses.Count ? addresses[i] : $"Area {i + 1}";
                viewModels.Add(new AreaViewModel(i + 1, addr, r, c));
            }

            AreasList.ItemsSource = viewModels;
            if (viewModels.Count > 0)
            {
                StatusText.Text = $"Detected {viewModels.Count} disjoint selection area(s).";
            }
            else
            {
                StatusText.Text = "No multi-area selection detected. Please select multiple ranges using Ctrl+Click in Excel.";
                CopyBtn.IsEnabled = false;
            }
        }

        private void CopyBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_areas == null || _areas.Count == 0) return;

            var mode = StackVertRadio.IsChecked == true
                ? MultiAreaCombineMode.StackVertically
                : MultiAreaCombineMode.PlaceSideBySide;

            var options = new MultiAreaCopyOptions { Mode = mode };
            var combined = MultiAreaRangeCopier.Combine(_areas, options);

            if (combined == null || combined.Length == 0)
            {
                _context.ShowMessage("No data available to combine.", "Combine Error", isError: true);
                return;
            }

            int totalRows = combined.GetUpperBound(0) - combined.GetLowerBound(0) + 1;
            int totalCols = combined.GetUpperBound(1) - combined.GetLowerBound(1) + 1;

            try
            {
                if (DestNewSheetRadio.IsChecked == true)
                {
                    _context.WriteMatrixToNewWorksheet("Combined_Ranges", combined);
                    _context.ShowMessage($"Combined {_areas.Count} areas into a new worksheet ({totalRows} rows × {totalCols} cols).", "Combine Complete");
                }
                else
                {
                    string destAddr = DestAddressBox.Text?.Trim() ?? "$E$1";
                    bool success = _context.WriteMatrixToAddress(destAddr, combined);
                    if (success)
                    {
                        _context.ShowMessage($"Combined {_areas.Count} areas into {destAddr} ({totalRows} rows × {totalCols} cols).", "Combine Complete");
                    }
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Failed to paste combined range: {ex.Message}", "Paste Error", isError: true);
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
