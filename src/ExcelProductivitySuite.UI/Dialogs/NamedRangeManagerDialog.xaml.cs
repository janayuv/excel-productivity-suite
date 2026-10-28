using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class NamedRangeManagerDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private List<NamedRangeInfo> _allNames = new List<NamedRangeInfo>();

        public class NamedRangeViewModel
        {
            public string Name { get; }
            public string RefersTo { get; }
            public string Scope { get; }
            public string Comment { get; }
            public bool IsValid { get; }
            public string TargetAddress { get; }
            public string StatusDisplay => IsValid ? "Valid" : "⚠️ Broken";

            public NamedRangeViewModel(NamedRangeInfo info)
            {
                Name = info.Name;
                RefersTo = info.RefersTo;
                Scope = info.Scope;
                Comment = info.Comment;
                IsValid = info.IsValid;
                TargetAddress = info.TargetAddress;
            }
        }

        public ObservableCollection<NamedRangeViewModel> FilteredNames { get; } = new ObservableCollection<NamedRangeViewModel>();

        public NamedRangeManagerDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            NamesGrid.ItemsSource = FilteredNames;
            LoadData();
        }

        private void LoadData()
        {
            _allNames = _context.GetNamedRanges().ToList();

            // Populate Scope Combo
            var scopes = new HashSet<string> { "All Scopes", "Workbook" };
            foreach (var n in _allNames)
            {
                if (!string.IsNullOrWhiteSpace(n.Scope) && n.Scope != "Workbook")
                {
                    scopes.Add(n.Scope);
                }
            }
            ScopeCombo.ItemsSource = scopes.ToList();
            ScopeCombo.SelectedIndex = 0;

            ApplyFilter();
        }

        private void ApplyFilter()
        {
            FilteredNames.Clear();
            string query = SearchBox.Text?.Trim().ToLowerInvariant() ?? "";
            string? scope = ScopeCombo.SelectedItem?.ToString();
            bool brokenOnly = BrokenOnlyCheck.IsChecked == true;

            foreach (var item in _allNames)
            {
                if (brokenOnly && item.IsValid) continue;

                if (!string.IsNullOrEmpty(scope) && scope != "All Scopes")
                {
                    if (!string.Equals(item.Scope, scope, StringComparison.OrdinalIgnoreCase)) continue;
                }

                if (!string.IsNullOrEmpty(query))
                {
                    if (!item.Name.ToLowerInvariant().Contains(query) &&
                        !item.RefersTo.ToLowerInvariant().Contains(query) &&
                        !item.Comment.ToLowerInvariant().Contains(query))
                    {
                        continue;
                    }
                }

                FilteredNames.Add(new NamedRangeViewModel(item));
            }

            int brokenCount = _allNames.Count(n => !n.IsValid);
            StatusText.Text = $"Showing {FilteredNames.Count} of {_allNames.Count} defined name(s). {brokenCount} broken reference(s) detected.";

            UpdateButtons();
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            ApplyFilter();
        }

        private void NamesGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateButtons();
        }

        private void UpdateButtons()
        {
            bool hasSelection = NamesGrid.SelectedItem is NamedRangeViewModel;
            GoToBtn.IsEnabled = hasSelection;
            DeleteBtn.IsEnabled = hasSelection;
            CopyAddrBtn.IsEnabled = hasSelection;
        }

        private void NamesGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            GoTo_Click(sender, e);
        }

        private void GoTo_Click(object sender, RoutedEventArgs e)
        {
            if (NamesGrid.SelectedItem is not NamedRangeViewModel selected) return;

            bool success = _context.GoToNamedRange(selected.Name);
            if (success)
            {
                // Range activated in Excel
                StatusText.Text = $"Navigated to '{selected.Name}' ({selected.RefersTo}).";
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (NamesGrid.SelectedItem is not NamedRangeViewModel selected) return;

            bool confirm = _context.Confirm($"Are you sure you want to delete named range '{selected.Name}'?", "Confirm Delete");
            if (!confirm) return;

            bool deleted = _context.DeleteNamedRange(selected.Name);
            if (deleted)
            {
                _allNames.RemoveAll(n => n.Name == selected.Name);
                ApplyFilter();
                _context.ShowMessage($"Named range '{selected.Name}' deleted.", "Deleted");
            }
        }

        private void CopyAddress_Click(object sender, RoutedEventArgs e)
        {
            if (NamesGrid.SelectedItem is not NamedRangeViewModel selected) return;

            try
            {
                Clipboard.SetText(selected.RefersTo);
                _context.ShowMessage($"Address '{selected.RefersTo}' copied to clipboard.", "Clipboard");
            }
            catch { }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
