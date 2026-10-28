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
    public partial class RemoveByPositionDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public RemoveByPositionDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private PositionRemovalOptions GetOptions()
        {
            var options = new PositionRemovalOptions();

            if (FromLeftRadio?.IsChecked == true)
            {
                options.Mode = PositionRemovalMode.FromLeft;
                if (int.TryParse(LeftCountTextBox?.Text, out int n)) options.Count = n;
            }
            else if (FromRightRadio?.IsChecked == true)
            {
                options.Mode = PositionRemovalMode.FromRight;
                if (int.TryParse(RightCountTextBox?.Text, out int n)) options.Count = n;
            }
            else if (FromMiddleRadio?.IsChecked == true)
            {
                options.Mode = PositionRemovalMode.FromMiddle;
                if (int.TryParse(MidStartTextBox?.Text, out int s)) options.StartIndex = s;
                if (int.TryParse(MidCountTextBox?.Text, out int c)) options.Count = c;
            }
            else if (DelimiterRadio?.IsChecked == true)
            {
                bool isBefore = (DelimModeComboBox?.SelectedItem as ComboBoxItem)?.Tag?.ToString() == "Before";
                options.Mode = isBefore ? PositionRemovalMode.BeforeDelimiter : PositionRemovalMode.AfterDelimiter;
                options.Delimiter = DelimiterTextBox?.Text ?? "-";
            }

            return options;
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            if (LeftCountTextBox != null && RightCountTextBox != null && MidStartTextBox != null &&
                MidCountTextBox != null && DelimModeComboBox != null && DelimiterTextBox != null)
            {
                LeftCountTextBox.IsEnabled = FromLeftRadio?.IsChecked == true;
                RightCountTextBox.IsEnabled = FromRightRadio?.IsChecked == true;
                MidStartTextBox.IsEnabled = FromMiddleRadio?.IsChecked == true;
                MidCountTextBox.IsEnabled = FromMiddleRadio?.IsChecked == true;
                DelimModeComboBox.IsEnabled = DelimiterRadio?.IsChecked == true;
                DelimiterTextBox.IsEnabled = DelimiterRadio?.IsChecked == true;
            }
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTable == null) return;

            var options = GetOptions();
            var previews = new List<PreviewItem>();

            for (int i = 0; i < _sampleOriginals.Count; i++)
            {
                string original = _sampleOriginals[i];
                string result = PositionRemover.Remove(original, options);

                previews.Add(new PreviewItem
                {
                    RowIndex = i + 1,
                    Original = original,
                    Result = result
                });
            }

            PreviewTable.ItemsSource = previews;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before removing characters.", "Sheet Protected", isError: true);
                return;
            }

            if (_context.IsMultiAreaSelection)
            {
                _context.ShowMessage("Please select a single contiguous range of cells to modify.", "Multi-Area Selection", isError: true);
                return;
            }

            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Remove Characters", isError: true);
                return;
            }

            var options = GetOptions();
            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int affected = 0;
            var updated = (object?[,])matrix.Clone();

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    var val = matrix[r, c];
                    if (val != null)
                    {
                        string str = val.ToString() ?? "";
                        if (!string.IsNullOrEmpty(str))
                        {
                            updated[r, c] = PositionRemover.Remove(str, options);
                            affected++;
                        }
                    }
                }
            }

            _context.SetSelectedValues(updated);
            DialogResult = true;
            Close();
            _context.ShowMessage($"Removed characters by position in {affected} cells.", "Remove Characters");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
