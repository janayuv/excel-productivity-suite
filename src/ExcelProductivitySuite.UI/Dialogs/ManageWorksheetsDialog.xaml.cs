using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class SheetVisibilityItem : INotifyPropertyChanged
    {
        private bool _isVisible;
        private bool _isVeryHidden;

        public string SheetName { get; set; } = "";

        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                _isVisible = value;
                if (_isVisible) _isVeryHidden = false;
                OnPropertyChanged(nameof(IsVisible));
                OnPropertyChanged(nameof(IsVeryHidden));
            }
        }

        public bool IsVeryHidden
        {
            get => _isVeryHidden;
            set
            {
                _isVeryHidden = value;
                if (_isVeryHidden) _isVisible = false;
                OnPropertyChanged(nameof(IsVeryHidden));
                OnPropertyChanged(nameof(IsVisible));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public partial class ManageWorksheetsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly ObservableCollection<SheetVisibilityItem> _items = new();

        public ManageWorksheetsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            LoadSheets();
        }

        private void LoadSheets()
        {
            _items.Clear();
            var infos = _context.GetWorksheetVisibilities();
            foreach (var info in infos)
            {
                _items.Add(new SheetVisibilityItem
                {
                    SheetName = info.SheetName,
                    IsVisible = info.IsVisible,
                    IsVeryHidden = info.IsVeryHidden
                });
            }

            SheetsGrid.ItemsSource = _items;
        }

        private void UnhideAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _items)
            {
                item.IsVisible = true;
            }
        }

        private void HideInactiveButton_Click(object sender, RoutedEventArgs e)
        {
            string? activeSheet = _context.ActiveWorksheetName;
            foreach (var item in _items)
            {
                if (!string.Equals(item.SheetName, activeSheet, StringComparison.OrdinalIgnoreCase))
                {
                    item.IsVisible = false;
                }
                else
                {
                    item.IsVisible = true;
                }
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            // Verify at least one sheet remains visible
            bool hasVisible = false;
            foreach (var item in _items)
            {
                if (item.IsVisible)
                {
                    hasVisible = true;
                    break;
                }
            }

            if (!hasVisible)
            {
                _context.ShowMessage("Excel requires at least one worksheet in the workbook to remain visible.", "Worksheet Visibility", isError: true);
                return;
            }

            foreach (var item in _items)
            {
                _context.SetWorksheetVisibility(item.SheetName, item.IsVisible, item.IsVeryHidden);
            }

            _context.ShowMessage("Worksheet visibility updated successfully.", "Worksheet Visibility");
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
