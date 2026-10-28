using System;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formatting;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class BatchNumberFormatDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private bool _isUpdatingMask = false;
        private double _activeSampleVal = 100.0;

        public BatchNumberFormatDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadActiveCellSample();
            UpdateGeneratedFormat();
        }

        private void LoadActiveCellSample()
        {
            var details = _context.GetActiveCellDetails();
            if (details != null && double.TryParse(details.ValueText, out double d))
            {
                _activeSampleVal = d;
            }
        }

        private void CategoryList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded) return;

            var item = CategoryList.SelectedItem as ListBoxItem;
            string tag = item?.Tag?.ToString() ?? "Currency";

            // Adjust controls visibility depending on category
            if (CurrencyRow != null)
            {
                CurrencyRow.Visibility = (tag == "Currency" || tag == "Accounting") ? Visibility.Visible : Visibility.Collapsed;
            }

            UpdateGeneratedFormat();
        }

        private void Control_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || _isUpdatingMask) return;
            UpdateGeneratedFormat();
        }

        private void UpdateGeneratedFormat()
        {
            if (!IsLoaded) return;

            var catItem = CategoryList.SelectedItem as ListBoxItem;
            string tag = catItem?.Tag?.ToString() ?? "Currency";

            var category = tag switch
            {
                "Currency" => NumberFormatCategory.Currency,
                "Accounting" => NumberFormatCategory.Accounting,
                "Percentage" => NumberFormatCategory.Percentage,
                "DecimalNumber" => NumberFormatCategory.DecimalNumber,
                "Integer" => NumberFormatCategory.Integer,
                "Scientific" => NumberFormatCategory.Scientific,
                "CompactUnits" => NumberFormatCategory.CompactUnits,
                _ => NumberFormatCategory.Currency
            };

            int decimals = DecimalsCombo.SelectedIndex;
            if (decimals < 0) decimals = 2;

            bool thousandSep = ThousandSepCheck.IsChecked == true;

            var negStyle = NegativeCombo.SelectedIndex switch
            {
                1 => NegativeNumberStyle.Parentheses,
                2 => NegativeNumberStyle.RedMinus,
                3 => NegativeNumberStyle.RedParentheses,
                _ => NegativeNumberStyle.StandardMinus
            };

            string currSymbol = "$";
            if (CurrencyCombo.SelectedItem is ComboBoxItem cItem && cItem.Tag != null)
            {
                currSymbol = cItem.Tag.ToString() ?? "$";
            }

            string mask = NumberFormatPresets.BuildFormat(category, decimals, thousandSep, negStyle, currSymbol);

            _isUpdatingMask = true;
            FormatMaskBox.Text = mask;
            _isUpdatingMask = false;

            UpdatePreviews(mask);
        }

        private void FormatMaskBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdatingMask || !IsLoaded) return;
            UpdatePreviews(FormatMaskBox.Text);
        }

        private void UpdatePreviews(string mask)
        {
            if (string.IsNullOrWhiteSpace(mask)) return;

            PosPreviewText.Text = NumberFormatPresets.FormatSamplePreview(12345.678, mask);
            NegPreviewText.Text = NumberFormatPresets.FormatSamplePreview(-12345.678, mask);
            ZeroPreviewText.Text = NumberFormatPresets.FormatSamplePreview(0.0, mask);
            ActiveCellPreviewText.Text = NumberFormatPresets.FormatSamplePreview(_activeSampleVal, mask);
        }

        private void ApplyBtn_Click(object sender, RoutedEventArgs e)
        {
            string mask = FormatMaskBox.Text;
            if (string.IsNullOrWhiteSpace(mask))
            {
                _context.ShowMessage("Please specify an Excel NumberFormat mask.", "Validation Error", isError: true);
                return;
            }

            try
            {
                _context.ApplyNumberFormat(mask, selectionOnly: true);
                _context.ShowMessage($"Successfully applied number format '{mask}' to active selection.", "Number Format Applied");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Failed to apply format: {ex.Message}", "Formatting Error", isError: true);
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
