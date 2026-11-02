using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class MathOperationsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleValues;

        public MathOperationsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            for (int i = 0; i <= 6; i++)
            {
                DecimalPlacesCombo.Items.Add(i.ToString());
            }
            DecimalPlacesCombo.SelectedIndex = 2;

            _sampleValues = _context.GetSampleOriginalValues(4);
            UpdatePreview();
        }

        private MathOperationType CurrentOperation
        {
            get
            {
                if (SubtractRadio?.IsChecked == true) return MathOperationType.Subtract;
                if (MultiplyRadio?.IsChecked == true) return MathOperationType.Multiply;
                if (DivideRadio?.IsChecked == true) return MathOperationType.Divide;
                if (PowerRadio?.IsChecked == true) return MathOperationType.Power;
                if (PercentRadio?.IsChecked == true) return MathOperationType.PercentChange;
                return MathOperationType.Add;
            }
        }

        private MathOperationOptions CurrentOptions
        {
            get
            {
                int decimals = 2;
                if (DecimalPlacesCombo?.SelectedItem != null)
                {
                    int.TryParse(DecimalPlacesCombo.SelectedItem.ToString(), out decimals);
                }

                return new MathOperationOptions
                {
                    RoundResults = RoundResultsCheck?.IsChecked == true,
                    DecimalPlaces = decimals,
                    TreatTextAsNumber = TreatTextCheck?.IsChecked == true,
                    SkipBlankCells = SkipBlanksCheck?.IsChecked == true
                };
            }
        }

        private bool TryGetOperand(out double operand)
        {
            operand = 0;
            string txt = OperandText?.Text?.Trim() ?? string.Empty;
            return double.TryParse(txt, NumberStyles.Any, CultureInfo.InvariantCulture, out operand) ||
                   double.TryParse(txt, NumberStyles.Any, CultureInfo.CurrentCulture, out operand);
        }

        private void Op_Changed(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void Operand_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewLabel == null || _sampleValues == null) return;

            if (!TryGetOperand(out double operand))
            {
                PreviewLabel.Text = "Please enter a valid numeric operand.";
                return;
            }

            if (_sampleValues.Count == 0)
            {
                // Fallback default samples
                var defaultSamples = new object?[4, 1] { { 100 }, { 250 }, { 15.5 }, { "80" } };
                var res = RangeMathOperations.Apply(defaultSamples, CurrentOperation, operand, CurrentOptions);
                var sb = new StringBuilder();
                for (int i = 0; i < 4; i++)
                {
                    sb.AppendLine($"{defaultSamples[i, 0]} -> {res.Matrix[i, 0]}");
                }
                PreviewLabel.Text = sb.ToString().TrimEnd();
                return;
            }

            var sampleMatrix = new object?[_sampleValues.Count, 1];
            for (int i = 0; i < _sampleValues.Count; i++)
            {
                sampleMatrix[i, 0] = _sampleValues[i];
            }

            var simResult = RangeMathOperations.Apply(sampleMatrix, CurrentOperation, operand, CurrentOptions);
            var previewSb = new StringBuilder();
            for (int i = 0; i < _sampleValues.Count; i++)
            {
                previewSb.AppendLine($"{_sampleValues[i]} -> {simResult.Matrix[i, 0]}");
            }
            PreviewLabel.Text = previewSb.ToString().TrimEnd();
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            if (!TryGetOperand(out double operand))
            {
                StatusLabel.Text = "Invalid operand. Please enter a valid number.";
                return;
            }

            try
            {
                object?[,]? matrix = _context.GetSelectedValues();
                if (matrix == null)
                {
                    StatusLabel.Text = "No range selected.";
                    return;
                }

                var result = RangeMathOperations.Apply(matrix, CurrentOperation, operand, CurrentOptions);
                _context.SetSelectedValues(result.Matrix);

                StatusLabel.Text = $"Calculated {result.ProcessedCount} cell(s).";
                _context.ShowMessage($"Successfully updated {result.ProcessedCount} cell(s). Skipped: {result.SkippedCount}, Errors: {result.ErrorCount}", "Math Operations");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error: {ex.Message}";
                _context.ShowMessage($"Math operation failed: {ex.Message}", "Math Operations", isError: true);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
