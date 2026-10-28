using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class PasteVisibleOnlyDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private object?[,]? _sourceData;

        public class MappingViewModel
        {
            public int SourceRowIndex { get; }
            public int TargetRowIndex { get; }
            public string ValuesDisplay { get; }

            public MappingViewModel(VisibleRowPasteMapping mapping)
            {
                SourceRowIndex = mapping.SourceRowIndex;
                TargetRowIndex = mapping.TargetRowIndex;
                ValuesDisplay = string.Join(" | ", mapping.Values ?? Array.Empty<object?>());
            }
        }

        public PasteVisibleOnlyDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadSourceData();
        }

        private void LoadSourceData()
        {
            // Try reading current selection values as source data
            _sourceData = _context.GetSelectedValues();
            if (_sourceData == null || _sourceData.Length == 0)
            {
                // Fallback default sample
                _sourceData = new object?[3, 2]
                {
                    { "Val_1A", "Val_1B" },
                    { "Val_2A", "Val_2B" },
                    { "Val_3A", "Val_3B" }
                };
            }

            TargetRowBox.Text = Math.Max(1, _context.SelectionStartRow).ToString();
            TargetColBox.Text = Math.Max(1, _context.SelectionStartColumn).ToString();

            RefreshMapping();
        }

        private void TargetCoordinates_Changed(object sender, TextChangedEventArgs e)
        {
            if (IsLoaded) RefreshMapping();
        }

        private void UseSelectionBtn_Click(object sender, RoutedEventArgs e)
        {
            TargetRowBox.Text = Math.Max(1, _context.SelectionStartRow).ToString();
            TargetColBox.Text = Math.Max(1, _context.SelectionStartColumn).ToString();
            RefreshMapping();
        }

        private void RefreshMapping()
        {
            if (_sourceData == null) return;

            if (!int.TryParse(TargetRowBox.Text, out int startRow) || startRow < 1) startRow = 1;
            if (!int.TryParse(TargetColBox.Text, out int startCol) || startCol < 1) startCol = 1;

            int srcRowCount = _sourceData.GetUpperBound(0) - _sourceData.GetLowerBound(0) + 1;
            var visibleRows = _context.GetVisibleRowIndices(startRow, srcRowCount);

            var rawMappings = VisibleCellsPasteEngine.ComputeMapping(_sourceData, visibleRows);
            var viewModels = new List<MappingViewModel>();

            foreach (var m in rawMappings)
            {
                viewModels.Add(new MappingViewModel(m));
            }

            MappingGrid.ItemsSource = viewModels;
            StatusText.Text = $"Source contains {srcRowCount} row(s). Found {visibleRows.Count} visible row(s) starting from row {startRow}.";
            PasteBtn.IsEnabled = viewModels.Count > 0;
        }

        private void PasteBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_sourceData == null) return;

            if (!int.TryParse(TargetRowBox.Text, out int startRow) || startRow < 1) startRow = 1;
            if (!int.TryParse(TargetColBox.Text, out int startCol) || startCol < 1) startCol = 1;

            try
            {
                bool success = _context.PasteIntoVisibleCells(_sourceData, startRow, startCol);
                if (success)
                {
                    _context.ShowMessage($"Successfully pasted into visible cells starting at row {startRow}, col {startCol}.", "Paste Complete");
                    DialogResult = true;
                    Close();
                }
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Paste failed: {ex.Message}", "Paste Error", isError: true);
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
