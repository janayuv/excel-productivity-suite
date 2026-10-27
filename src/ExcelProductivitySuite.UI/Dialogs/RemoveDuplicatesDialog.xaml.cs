using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class ColumnSelectionItem : INotifyPropertyChanged
    {
        public int ColumnIndex { get; set; }
        public string HeaderName { get; set; } = "";
        public string DefaultName { get; set; } = "";

        private string _displayName = "";
        public string DisplayName
        {
            get => _displayName;
            set
            {
                if (_displayName != value)
                {
                    _displayName = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
                }
            }
        }

        private bool _isSelected = true;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    public partial class RemoveDuplicatesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;
        private readonly ObservableCollection<ColumnSelectionItem> _columns = new ObservableCollection<ColumnSelectionItem>();
        private int _detectedDuplicatesCount = 0;

        public RemoveDuplicatesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeInfoText.Text = $"Range: {_context.SelectionAddress ?? "No Selection"}";
            _matrix = _context.GetSelectedValues();

            InitializeColumns();
            ColumnsList.ItemsSource = _columns;
            UpdateDuplicateCalculation();
        }

        private void InitializeColumns()
        {
            _columns.Clear();
            if (_matrix == null) return;

            int cStart = _matrix.GetLowerBound(1);
            int cEnd = _matrix.GetUpperBound(1);
            int rStart = _matrix.GetLowerBound(0);

            int startCol = _context.SelectionStartColumn;

            for (int c = cStart; c <= cEnd; c++)
            {
                int offset = c - cStart;
                string colLetter = GetColumnLetter(startCol + offset);
                string defaultName = $"Column {colLetter}";
                string headerName = _matrix[rStart, c]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(headerName)) headerName = defaultName;

                var item = new ColumnSelectionItem
                {
                    ColumnIndex = offset,
                    DefaultName = defaultName,
                    HeaderName = $"{colLetter}: {headerName}",
                    DisplayName = defaultName,
                    IsSelected = true
                };

                _columns.Add(item);
            }

            UpdateColumnLabels();
        }

        private void UpdateColumnLabels()
        {
            bool hasHeaders = HasHeadersCheckBox?.IsChecked == true;
            foreach (var col in _columns)
            {
                col.DisplayName = hasHeaders ? col.HeaderName : col.DefaultName;
            }
        }

        private static string GetColumnLetter(int columnNumber)
        {
            string colLetter = "";
            while (columnNumber > 0)
            {
                int mod = (columnNumber - 1) % 26;
                colLetter = (char)('A' + mod) + colLetter;
                columnNumber = (columnNumber - mod) / 26;
            }
            return string.IsNullOrEmpty(colLetter) ? "A" : colLetter;
        }

        private DuplicateDetectionOptions GetOptions()
        {
            var selectedIndices = _columns
                .Where(c => c.IsSelected)
                .Select(c => c.ColumnIndex)
                .ToArray();

            return new DuplicateDetectionOptions
            {
                KeepStrategy = KeepFirstRadio.IsChecked == true ? DuplicateKeepStrategy.KeepFirst : DuplicateKeepStrategy.KeepLast,
                MatchCase = MatchCaseCheckBox.IsChecked == true,
                HasHeaders = HasHeadersCheckBox.IsChecked == true,
                CompareEntireRows = selectedIndices.Length == 0 || selectedIndices.Length == _columns.Count,
                KeyColumnIndices = selectedIndices.Length > 0 ? selectedIndices : null
            };
        }

        private void UpdateDuplicateCalculation()
        {
            if (_matrix == null || DuplicateSummaryText == null) return;

            var options = GetOptions();
            bool[] isDuplicate = DuplicateFinder.IdentifyDuplicateRows(_matrix, options);

            int dupCount = 0;
            foreach (bool dup in isDuplicate)
            {
                if (dup) dupCount++;
            }

            _detectedDuplicatesCount = dupCount;
            int totalRows = _matrix.GetLength(0);
            int dataRows = options.HasHeaders ? Math.Max(0, totalRows - 1) : totalRows;

            if (dupCount == 0)
            {
                DuplicateSummaryText.Text = $"0 duplicates found across {dataRows} rows.";
            }
            else
            {
                DuplicateSummaryText.Text = $"{dupCount} duplicate rows detected out of {dataRows} data rows.";
            }
        }

        private void CriteriaChanged(object sender, RoutedEventArgs e)
        {
            UpdateDuplicateCalculation();
        }

        private void HeadersCheckboxChanged(object sender, RoutedEventArgs e)
        {
            UpdateColumnLabels();
            UpdateDuplicateCalculation();
        }

        private void ColumnCheckChanged(object sender, RoutedEventArgs e)
        {
            UpdateDuplicateCalculation();
        }

        private void SelectAllColumns_Click(object sender, RoutedEventArgs e)
        {
            foreach (var col in _columns) col.IsSelected = true;
            UpdateDuplicateCalculation();
        }

        private void UnselectAllColumns_Click(object sender, RoutedEventArgs e)
        {
            foreach (var col in _columns) col.IsSelected = false;
            UpdateDuplicateCalculation();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before removing duplicates.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells to remove duplicates.", "Remove Duplicates", isError: true);
                return;
            }

            if (_matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Remove Duplicates", isError: true);
                return;
            }

            var options = GetOptions();
            bool[] isDuplicate = DuplicateFinder.IdentifyDuplicateRows(_matrix, options);

            int dupCount = 0;
            foreach (bool dup in isDuplicate)
            {
                if (dup) dupCount++;
            }

            if (dupCount == 0)
            {
                _context.ShowMessage("No duplicate rows were found in the selected range.", "Remove Duplicates");
                DialogResult = true;
                Close();
                return;
            }

            bool proceed = _context.Confirm($"Are you sure you want to clear {dupCount} duplicate rows in the selection? This operation preserves row layout and non-duplicate records.", "Confirm Duplicate Removal");
            if (!proceed) return;

            int rStart = _matrix.GetLowerBound(0);
            int rEnd = _matrix.GetUpperBound(0);
            int cStart = _matrix.GetLowerBound(1);
            int cEnd = _matrix.GetUpperBound(1);

            var updated = (object?[,])_matrix.Clone();

            for (int r = rStart; r <= rEnd; r++)
            {
                int rowIndex = r - rStart;
                if (isDuplicate[rowIndex])
                {
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        updated[r, c] = "";
                    }
                }
            }

            _context.SetSelectedValues(updated);
            _context.ShowMessage($"Successfully cleared {dupCount} duplicate rows.", "Remove Duplicates");

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
