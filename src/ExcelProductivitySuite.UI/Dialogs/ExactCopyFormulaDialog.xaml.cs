using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formulas;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ExactCopyFormulaDialog : BaseDialogWindow
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        private readonly IExcelContext _context;

        public TextBox SourceTextBox => SourceRangeText;
        public TextBox DestinationTextBox => DestinationRangeText;
        public Button SourceSelectButton => SelectSourceRangeButton;
        public Button DestinationSelectButton => SelectDestinationRangeButton;
        public CheckBox CopyFormattingCheckBox => CopyFormattingCheck;

        public ExactCopyFormulaDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SourceRangeText.Text = _context.SelectionAddress ?? string.Empty;
        }

        private void SelectSourceRange_Click(object sender, RoutedEventArgs e)
        {
            PromptRangeForTextBox(SourceRangeText, "Ranges to be copied :");
        }

        private void SelectDestinationRange_Click(object sender, RoutedEventArgs e)
        {
            PromptRangeForTextBox(DestinationRangeText, "Please select a cell or range to paste :");
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

                string? selected = _context.PromptForRangeAddress(prompt, "Exact Formula Copy", current);
                if (!string.IsNullOrWhiteSpace(selected))
                {
                    targetBox.Text = selected;
                    StatusLabel.Text = $"Selected range: {selected}";
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

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            string srcAddr = SourceRangeText.Text?.Trim() ?? string.Empty;
            string dstAddr = DestinationRangeText.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(srcAddr))
            {
                StatusLabel.Text = "Please enter or select a source range.";
                return;
            }

            if (string.IsNullOrWhiteSpace(dstAddr))
            {
                StatusLabel.Text = "Please enter a destination range or top-left cell.";
                return;
            }

            try
            {
                object?[,]? sourceFormulas = _context.GetRangeFormulasByAddress(srcAddr);
                if (sourceFormulas == null)
                {
                    StatusLabel.Text = "Unable to read formulas from source range.";
                    return;
                }

                int sRowStart = sourceFormulas.GetLowerBound(0);
                int sRowEnd = sourceFormulas.GetUpperBound(0);
                int sColStart = sourceFormulas.GetLowerBound(1);
                int sColEnd = sourceFormulas.GetUpperBound(1);

                int rows = sRowEnd - sRowStart + 1;
                int cols = sColEnd - sColStart + 1;

                object?[,] target = FormulaExactCopier.CopyExactMatrix(sourceFormulas, rows, cols);

                if (ConfirmOverwriteCheck.IsChecked == true)
                {
                    object?[,]? existingValues = _context.GetRangeValuesByAddress(dstAddr);
                    bool hasData = false;
                    if (existingValues != null)
                    {
                        foreach (var val in existingValues)
                        {
                            if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
                            {
                                hasData = true;
                                break;
                            }
                        }
                    }

                    if (hasData)
                    {
                        if (!_context.Confirm($"The destination '{dstAddr}' contains existing data. Do you want to overwrite it?", "Confirm Overwrite"))
                        {
                            StatusLabel.Text = "Operation cancelled.";
                            return;
                        }
                    }
                }

                _context.SetRangeFormulasByAddress(dstAddr, target);

                if (CopyFormattingCheck?.IsChecked == true)
                {
                    _context.CopyRangeFormatsByAddress(srcAddr, dstAddr);
                }

                _context.ShowMessage($"Successfully copied {rows * cols} formula(s) verbatim to '{dstAddr}'.", "Exact Copy Formulas");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error: {ex.Message}";
                _context.ShowMessage($"Failed to copy formulas: {ex.Message}", "Exact Copy Error", isError: true);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
