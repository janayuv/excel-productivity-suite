using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class BatchUnprotectSheetsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private bool _isPasswordVisible = false;

        public class SheetItem : INotifyPropertyChanged
        {
            private bool _isSelected;
            public string SheetName { get; }
            public bool IsProtected { get; }

            public string DisplayName => IsProtected ? $"{SheetName} (🔒 Protected)" : SheetName;

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

            public SheetItem(string name, bool isProtected, bool isSelected = true)
            {
                SheetName = name;
                IsProtected = isProtected;
                _isSelected = isSelected;
            }

            public event PropertyChangedEventHandler? PropertyChanged;
        }

        public ObservableCollection<SheetItem> SheetItems { get; } = new ObservableCollection<SheetItem>();

        public BatchUnprotectSheetsDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadSheets();
            SheetsListBox.ItemsSource = SheetItems;
        }

        private void LoadSheets()
        {
            SheetItems.Clear();
            var statuses = _context.GetSheetProtectionStatuses();
            if (statuses.Count > 0)
            {
                foreach (var s in statuses)
                {
                    // By default select protected sheets
                    SheetItems.Add(new SheetItem(s.SheetName, s.IsProtected, s.IsProtected));
                }
            }
            else
            {
                var names = _context.GetWorksheetNames();
                foreach (var n in names)
                {
                    SheetItems.Add(new SheetItem(n, false, true));
                }
            }

            UpdateStatus();
        }

        private void UpdateStatus()
        {
            int selectedCount = SheetItems.Count(s => s.IsSelected);
            StatusText.Text = $"{selectedCount} of {SheetItems.Count} worksheet(s) selected.";
        }

        private void SelectProtected_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in SheetItems) s.IsSelected = s.IsProtected;
            UpdateStatus();
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in SheetItems) s.IsSelected = true;
            UpdateStatus();
        }

        private void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in SheetItems) s.IsSelected = false;
            UpdateStatus();
        }

        private void TogglePassword_Click(object sender, RoutedEventArgs e)
        {
            _isPasswordVisible = !_isPasswordVisible;
            if (_isPasswordVisible)
            {
                PasswordVisibleBox.Text = PasswordBox.Password;
                PasswordBox.Visibility = Visibility.Collapsed;
                PasswordVisibleBox.Visibility = Visibility.Visible;
                TogglePwdBtn.Content = "🙈";
            }
            else
            {
                PasswordBox.Password = PasswordVisibleBox.Text;
                PasswordVisibleBox.Visibility = Visibility.Collapsed;
                PasswordBox.Visibility = Visibility.Visible;
                TogglePwdBtn.Content = "👁";
            }
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (!_isPasswordVisible)
            {
                PasswordVisibleBox.Text = PasswordBox.Password;
            }
        }

        private void PasswordVisibleBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (_isPasswordVisible)
            {
                PasswordBox.Password = PasswordVisibleBox.Text;
            }
        }

        private void Unprotect_Click(object sender, RoutedEventArgs e)
        {
            var selectedSheets = SheetItems.Where(s => s.IsSelected).Select(s => s.SheetName).ToList();
            if (selectedSheets.Count == 0)
            {
                _context.ShowMessage("Please select at least one worksheet to unprotect.", "Selection Required", isError: true);
                return;
            }

            string pwd = _isPasswordVisible ? PasswordVisibleBox.Text : PasswordBox.Password;

            var (successCount, errors) = _context.BatchUnprotectWorksheets(selectedSheets, string.IsNullOrEmpty(pwd) ? null : pwd);

            if (errors.Count == 0)
            {
                _context.ShowMessage($"Successfully unprotected {successCount} worksheet(s).", "Worksheets Unprotected");
                DialogResult = true;
                Close();
            }
            else
            {
                string msg = $"Unprotected {successCount} sheet(s) with {errors.Count} error(s) (incorrect password or error):\n" + string.Join("\n", errors.Take(5));
                _context.ShowMessage(msg, "Unprotection Notice", isError: true);
                DialogResult = true;
                Close();
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
