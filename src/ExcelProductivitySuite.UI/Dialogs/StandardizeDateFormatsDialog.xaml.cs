using System;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formatting;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class StandardizeDateFormatsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public StandardizeDateFormatsDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadStandards();
        }

        private void LoadStandards()
        {
            var standards = DateFormatStandardizer.GetStandards();
            StandardsList.ItemsSource = standards;
            if (standards.Count > 0)
            {
                StandardsList.SelectedIndex = 0;
            }
        }

        private void StandardsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StandardsList.SelectedItem is DateFormatStandardItem item)
            {
                CustomMaskBox.Text = item.ExcelFormat;
                StatusText.Text = $"Selected '{item.Name}'. Output example: {item.SampleOutput}";
            }
        }

        private void ApplyBtn_Click(object sender, RoutedEventArgs e)
        {
            string mask = CustomMaskBox.Text;
            if (string.IsNullOrWhiteSpace(mask))
            {
                _context.ShowMessage("Please specify an Excel date format mask.", "Validation Error", isError: true);
                return;
            }

            bool convertText = ConvertTextDatesCheck.IsChecked == true;

            try
            {
                int textConverted = 0;
                if (convertText)
                {
                    object?[,]? rawVals = _context.GetSelectedValues();
                    if (rawVals != null && rawVals.Length > 0)
                    {
                        var result = DateFormatStandardizer.StandardizeRange(rawVals, mask, convertTextDates: true);
                        if (result.TextDatesConverted > 0)
                        {
                            _context.SetSelectedValues(result.ProcessedValues);
                            textConverted = result.TextDatesConverted;
                        }
                    }
                }

                _context.ApplyNumberFormat(mask, selectionOnly: true);

                string msg = textConverted > 0
                    ? $"Applied date format '{mask}' to selection. Converted {textConverted} text date(s) to true Excel dates."
                    : $"Applied date format '{mask}' to selection.";

                _context.ShowMessage(msg, "Date Format Applied");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Failed to apply date format: {ex.Message}", "Formatting Error", isError: true);
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
