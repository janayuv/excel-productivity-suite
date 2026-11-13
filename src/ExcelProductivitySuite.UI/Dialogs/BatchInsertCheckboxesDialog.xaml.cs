using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class BatchInsertCheckboxesDialog : BaseDialogWindow
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        private readonly IExcelContext _context;

        public TextBox SourceTextBox => SourceRangeText;
        public TextBox TargetTextBox => TargetRangeText;
        public RadioButton RadioNoneOption => RadioNone;
        public RadioButton RadioOwnCellOption => RadioOwnCell;
        public RadioButton RadioOtherCellsOption => RadioOtherCells;
        public CheckBox UseLabelCheckBox => UseLabelCheck;
        public CheckBox DisplayLogicalValueCheckBox => DisplayLogicalValueCheck;
        public CheckBox InitialCheckedCheckBox => InitialCheckedCheck;

        public BatchInsertCheckboxesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SourceRangeText.Text = _context.SelectionAddress ?? "$A$1";
            UpdateRangeCountDisplay();
        }

        private void SourceRangeText_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateRangeCountDisplay();
        }

        private void UpdateRangeCountDisplay()
        {
            if (RangeCountText == null) return;
            string addr = SourceRangeText?.Text?.Trim() ?? string.Empty;
            if (CheckboxEngine.TryParseRangeAddress(addr, out _, out int sr, out int sc, out int er, out int ec))
            {
                int count = (er - sr + 1) * (ec - sc + 1);
                RangeCountText.Text = $"Cells selected: {count} ({CheckboxEngine.ColumnIndexToLetter(sc)}{sr}:{CheckboxEngine.ColumnIndexToLetter(ec)}{er})";
            }
            else
            {
                RangeCountText.Text = "Please enter a valid range address (e.g. $A$2:$A$10).";
            }
        }

        private void LinkMode_Changed(object sender, RoutedEventArgs e)
        {
            if (OtherCellsPanel == null) return;
            bool isOtherCells = RadioOtherCells?.IsChecked == true;
            OtherCellsPanel.IsEnabled = isOtherCells;
            OtherCellsPanel.Opacity = isOtherCells ? 1.0 : 0.6;
        }

        private void SelectSourceRange_Click(object sender, RoutedEventArgs e)
        {
            PromptRangeForTextBox(SourceRangeText, "Select cell range to insert checkboxes:");
        }

        private void SelectTargetRange_Click(object sender, RoutedEventArgs e)
        {
            PromptRangeForTextBox(TargetRangeText, "Select target cell or range for linked TRUE/FALSE values:");
        }

        public void PromptRangeForTextBox(TextBox targetBox, string prompt)
        {
            if (targetBox == null) return;

            string current = targetBox.Text?.Trim() ?? string.Empty;
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

                string? selected = _context.PromptForRangeAddress(prompt, "Batch Insert Smart Checkboxes", current);
                if (!string.IsNullOrWhiteSpace(selected))
                {
                    targetBox.Text = selected;
                    StatusLabel.Text = $"Selected: {selected}";
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
                targetBox.Focus();
            }
        }

        private void DeleteCheckboxesLink_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var deleteDlg = new BatchDeleteCheckboxesDialog(_context);
                deleteDlg.ShowModal(_context);
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error opening delete tool: {ex.Message}";
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            string sourceAddr = SourceRangeText.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(sourceAddr))
            {
                StatusLabel.Text = "Please enter or select a valid target cell range.";
                SourceRangeText.Focus();
                return;
            }

            var linkMode = CheckboxLinkMode.OwnCell;
            if (RadioNone.IsChecked == true)
            {
                linkMode = CheckboxLinkMode.None;
            }
            else if (RadioOtherCells.IsChecked == true)
            {
                linkMode = CheckboxLinkMode.OtherCells;
            }

            string? targetAddr = TargetRangeText.Text?.Trim();
            if (linkMode == CheckboxLinkMode.OtherCells && string.IsNullOrWhiteSpace(targetAddr))
            {
                StatusLabel.Text = "Please specify a target cell or range to link checkboxes to.";
                TargetRangeText.Focus();
                return;
            }

            var options = new BatchInsertCheckboxOptions
            {
                SourceRangeAddress = sourceAddr,
                LinkMode = linkMode,
                TargetLinkRangeAddress = targetAddr,
                UseCellContentAsLabel = UseLabelCheck.IsChecked == true,
                DisplayLogicalValue = DisplayLogicalValueCheck.IsChecked == true,
                InitialChecked = InitialCheckedCheck.IsChecked == true
            };

            try
            {
                var result = _context.InsertCheckboxes(options);
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
                StatusLabel.Text = $"Insertion failed: {ex.Message}";
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
