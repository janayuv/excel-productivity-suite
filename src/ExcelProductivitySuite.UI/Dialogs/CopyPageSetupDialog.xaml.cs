using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class SheetSelectionItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        public string Name { get; set; } = string.Empty;

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

    public partial class CopyPageSetupDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly ObservableCollection<SheetSelectionItem> _targetSheets = new();

        public CopyPageSetupDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            TargetSheetsList.ItemsSource = _targetSheets;
            PopulateSheets();
        }

        private void PopulateSheets()
        {
            SourceSheetCombo.Items.Clear();
            var allSheets = _context.GetWorksheetNames();
            string activeSheet = _context.ActiveWorksheetName ?? string.Empty;

            int activeIndex = 0;
            for (int i = 0; i < allSheets.Count; i++)
            {
                SourceSheetCombo.Items.Add(allSheets[i]);
                if (string.Equals(allSheets[i], activeSheet, StringComparison.OrdinalIgnoreCase))
                {
                    activeIndex = i;
                }
            }

            if (SourceSheetCombo.Items.Count > 0)
            {
                SourceSheetCombo.SelectedIndex = activeIndex;
            }

            UpdateTargetSheets();
        }

        private void SourceSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateTargetSheets();
        }

        private void UpdateTargetSheets()
        {
            _targetSheets.Clear();
            string selectedSource = SourceSheetCombo?.SelectedItem?.ToString() ?? string.Empty;
            var allSheets = _context.GetWorksheetNames();

            foreach (var s in allSheets)
            {
                if (!string.Equals(s, selectedSource, StringComparison.OrdinalIgnoreCase))
                {
                    _targetSheets.Add(new SheetSelectionItem { Name = s, IsSelected = true });
                }
            }
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _targetSheets)
            {
                item.IsSelected = true;
            }
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _targetSheets)
            {
                item.IsSelected = false;
            }
        }

        private void CopySetup_Click(object sender, RoutedEventArgs e)
        {
            string sourceSheet = SourceSheetCombo?.SelectedItem?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(sourceSheet))
            {
                StatusLabel.Text = "Please select a source worksheet.";
                return;
            }

            var targets = _targetSheets.Where(s => s.IsSelected).Select(s => s.Name).ToList();
            if (targets.Count == 0)
            {
                StatusLabel.Text = "Please select at least one target worksheet.";
                return;
            }

            var options = new PageSetupOptions
            {
                CopyOrientationAndSize = OrientationCheck.IsChecked == true,
                CopyMargins = MarginsCheck.IsChecked == true,
                CopyPrintTitles = PrintTitlesCheck.IsChecked == true,
                CopyScaling = ScalingCheck.IsChecked == true,
                CopyHeadersFooters = HeadersCheck.IsChecked == true,
                CopySheetOptions = SheetOptionsCheck.IsChecked == true
            };

            try
            {
                int copied = _context.CopyPageSetup(sourceSheet, targets, options);
                StatusLabel.Text = $"Successfully copied page setup to {copied} worksheet(s).";
                _context.ShowMessage($"Copied page setup from '{sourceSheet}' to {copied} worksheet(s).", "Copy Page Setup");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error: {ex.Message}";
                _context.ShowMessage($"Failed to copy page setup: {ex.Message}", "Page Setup Error", isError: true);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
