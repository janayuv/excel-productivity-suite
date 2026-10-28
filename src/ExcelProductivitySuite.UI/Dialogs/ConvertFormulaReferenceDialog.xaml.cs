using System;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formulas;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ConvertFormulaReferenceDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public ConvertFormulaReferenceDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeAddressLabel.Text = _context.SelectionAddress ?? "Selection";

            // Extract first formula from selection if available to populate the sample box
            var formulas = _context.GetSelectedFormulas();
            if (formulas != null)
            {
                int rStart = formulas.GetLowerBound(0);
                int rEnd = formulas.GetUpperBound(0);
                int cStart = formulas.GetLowerBound(1);
                int cEnd = formulas.GetUpperBound(1);

                for (int r = rStart; r <= rEnd; r++)
                {
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        if (formulas[r, c] is string f && f.StartsWith("="))
                        {
                            SampleFormulaText.Text = f;
                            break;
                        }
                    }
                }
            }

            UpdatePreview();
        }

        private ReferenceConversionMode SelectedMode
        {
            get
            {
                if (RowAbsoluteRadio?.IsChecked == true) return ReferenceConversionMode.ToRowAbsolute;
                if (ColAbsoluteRadio?.IsChecked == true) return ReferenceConversionMode.ToColumnAbsolute;
                if (RelativeRadio?.IsChecked == true) return ReferenceConversionMode.ToRelative;
                return ReferenceConversionMode.ToAbsolute;
            }
        }

        private void Mode_Changed(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void Sample_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewResultText == null || SampleFormulaText == null) return;

            string sample = SampleFormulaText.Text;
            if (string.IsNullOrWhiteSpace(sample))
            {
                PreviewResultText.Text = string.Empty;
                return;
            }

            try
            {
                PreviewResultText.Text = FormulaReferenceConverter.ConvertFormulaReferences(sample, SelectedMode);
            }
            catch
            {
                PreviewResultText.Text = sample;
            }
        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var formulas = _context.GetSelectedFormulas();
                if (formulas == null)
                {
                    StatusLabel.Text = "No selection or unable to read formulas.";
                    return;
                }

                var (resultMatrix, modified) = FormulaReferenceConverter.ConvertMatrix(formulas, SelectedMode);
                _context.SetSelectedFormulas(resultMatrix);

                StatusLabel.Text = $"Successfully updated {modified} formula references.";
                _context.ShowMessage($"Converted reference styles in {modified} formula cell(s).", "Convert Formula References");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error: {ex.Message}";
                _context.ShowMessage($"Failed to convert references: {ex.Message}", "Conversion Error", isError: true);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
