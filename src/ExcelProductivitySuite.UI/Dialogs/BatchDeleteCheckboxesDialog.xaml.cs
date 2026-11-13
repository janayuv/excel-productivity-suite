using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class BatchDeleteCheckboxesDialog : BaseDialogWindow
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        private readonly IExcelContext _context;

        public TextBox RangeBox => RangeTextBox;
        public RadioButton RadioSelectionScope => RadioSelection;
        public RadioButton RadioSheetScope => RadioSheet;

        public BatchDeleteCheckboxesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeTextBox.Text = _context.SelectionAddress ?? "$A$1";
            UpdateCountSummary();
        }

        private void Scope_Changed(object sender, RoutedEventArgs e)
        {
            if (RangeSelectionPanel == null) return;
            bool isSelection = RadioSelection?.IsChecked == true;
            RangeSelectionPanel.IsEnabled = isSelection;
            RangeSelectionPanel.Opacity = isSelection ? 1.0 : 0.6;
            UpdateCountSummary();
        }

        private void UpdateCountSummary()
        {
            if (CountSummaryText == null) return;
            try
            {
                if (RadioSheet?.IsChecked == true)
                {
                    int total = _context.GetCheckboxCount(null);
                    CountSummaryText.Text = $"Detected {total} checkbox(es) on the entire worksheet.";
                }
                else
                {
                    string addr = RangeTextBox?.Text?.Trim() ?? string.Empty;
                    int count = _context.GetCheckboxCount(addr);
                    CountSummaryText.Text = $"Detected {count} checkbox(es) in the specified range.";
                }
            }
            catch
            {
                CountSummaryText.Text = "Unable to inspect checkboxes.";
            }
        }

        private void SelectRange_Click(object sender, RoutedEventArgs e)
        {
            string current = RangeTextBox.Text?.Trim() ?? string.Empty;
            IntPtr hwndOwner = IntPtr.Zero;

            try
            {
                hwndOwner = _context.MainWindowHandle;
            }
            catch
            {
                hwndOwner = IntPtr.Zero;
            }

            try
            {
                Visibility = Visibility.Hidden;
                if (hwndOwner != IntPtr.Zero)
                {
                    EnableWindow(hwndOwner, true);
                }

                string? selected = _context.PromptForRangeAddress("Select range to delete checkboxes from:", "Batch Delete Checkboxes", current);
                if (!string.IsNullOrWhiteSpace(selected))
                {
                    RangeTextBox.Text = selected;
                    UpdateCountSummary();
                }
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Selection error: {ex.Message}";
            }
            finally
            {
                if (hwndOwner != IntPtr.Zero)
                {
                    EnableWindow(hwndOwner, false);
                }
                Visibility = Visibility.Visible;
                Activate();
                RangeTextBox.Focus();
            }
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            var options = new BatchDeleteCheckboxOptions
            {
                Scope = RadioSheet.IsChecked == true ? CheckboxDeleteScope.ActiveWorksheet : CheckboxDeleteScope.Selection,
                RangeAddress = RangeTextBox.Text?.Trim()
            };

            try
            {
                var result = _context.DeleteCheckboxes(options);
                if (result.Success)
                {
                    DialogResult = true;
                    Close();
                }
                else
                {
                    StatusLabel.Text = result.Message;
                }
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Deletion failed: {ex.Message}";
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
