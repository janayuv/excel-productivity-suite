using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ColumnNavigatorDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private List<WorksheetColumnInfo> _allColumns = new List<WorksheetColumnInfo>();

        public class ColumnViewModel
        {
            public int Index { get; }
            public string ColumnLetter { get; }
            public string Header { get; }
            public double Width { get; }
            public bool IsHidden { get; }
            public string VisibilityDisplay => IsHidden ? "Hidden" : "Visible";

            public ColumnViewModel(WorksheetColumnInfo info)
            {
                Index = info.Index;
                ColumnLetter = info.ColumnLetter;
                Header = string.IsNullOrWhiteSpace(info.Header) ? "(Empty Header)" : info.Header;
                Width = info.Width;
                IsHidden = info.IsHidden;
            }
        }

        public ObservableCollection<ColumnViewModel> FilteredColumns { get; } = new ObservableCollection<ColumnViewModel>();

        public ColumnNavigatorDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            ColumnsGrid.ItemsSource = FilteredColumns;
            LoadColumns();
        }

        private void LoadColumns()
        {
            _allColumns = _context.GetWorksheetColumns(maxCols: 200).ToList();
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (_allColumns == null || FilterCombo == null || ColumnsGrid == null || StatusText == null) return;
            FilteredColumns.Clear();
            string query = SearchBox?.Text?.Trim().ToLowerInvariant() ?? "";
            string filterMode = ((System.Windows.Controls.ComboBoxItem)FilterCombo.SelectedItem)?.Content?.ToString() ?? "All Columns";

            foreach (var col in _allColumns)
            {
                if (filterMode == "Visible Only" && col.IsHidden) continue;
                if (filterMode == "Hidden Only" && !col.IsHidden) continue;

                if (!string.IsNullOrEmpty(query))
                {
                    if (!col.ColumnLetter.ToLowerInvariant().Contains(query) &&
                        !col.Header.ToLowerInvariant().Contains(query))
                    {
                        continue;
                    }
                }

                FilteredColumns.Add(new ColumnViewModel(col));
            }

            int hiddenCount = _allColumns.Count(c => c.IsHidden);
            StatusText.Text = $"Showing {FilteredColumns.Count} of {_allColumns.Count} column(s). {hiddenCount} column(s) currently hidden.";
            UpdateButtons();
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            ApplyFilter();
        }

        private void ColumnsGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateButtons();
        }

        private void UpdateButtons()
        {
            var selectedItems = ColumnsGrid.SelectedItems.Cast<ColumnViewModel>().ToList();
            bool hasSelection = selectedItems.Count > 0;

            JumpBtn.IsEnabled = hasSelection;
            HideBtn.IsEnabled = hasSelection;
            UnhideBtn.IsEnabled = hasSelection;
            AutoFitBtn.IsEnabled = hasSelection;
        }

        private void ColumnsGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Jump_Click(sender, e);
        }

        private void Jump_Click(object sender, RoutedEventArgs e)
        {
            if (ColumnsGrid.SelectedItem is not ColumnViewModel selected) return;
            _context.SelectColumn(selected.Index);
            StatusText.Text = $"Selected Column {selected.ColumnLetter} ({selected.Header}).";
        }

        private void Hide_Click(object sender, RoutedEventArgs e)
        {
            var selectedIndices = ColumnsGrid.SelectedItems.Cast<ColumnViewModel>().Select(c => c.Index).ToList();
            if (selectedIndices.Count == 0) return;

            _context.SetColumnsVisibility(selectedIndices, true);
            LoadColumns();
        }

        private void Unhide_Click(object sender, RoutedEventArgs e)
        {
            var selectedIndices = ColumnsGrid.SelectedItems.Cast<ColumnViewModel>().Select(c => c.Index).ToList();
            if (selectedIndices.Count == 0) return;

            _context.SetColumnsVisibility(selectedIndices, false);
            LoadColumns();
        }

        private void AutoFit_Click(object sender, RoutedEventArgs e)
        {
            var selectedIndices = ColumnsGrid.SelectedItems.Cast<ColumnViewModel>().Select(c => c.Index).ToList();
            if (selectedIndices.Count == 0) return;

            _context.AutoFitColumns(selectedIndices);
            LoadColumns();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
