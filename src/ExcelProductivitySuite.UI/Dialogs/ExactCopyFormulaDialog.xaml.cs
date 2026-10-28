using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formulas;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ExactCopyFormulaDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public ExactCopyFormulaDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SourceRangeText.Text = _context.SelectionAddress ?? string.Empty;
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
