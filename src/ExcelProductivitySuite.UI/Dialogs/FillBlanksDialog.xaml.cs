using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FillBlanksDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public FillBlanksDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();
        }

        private void StrategyChanged(object sender, RoutedEventArgs e)
        {
            if (FixedValueTextBox != null)
            {
                FixedValueTextBox.IsEnabled = FillFixedRadio.IsChecked == true;
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Fill Blank Cells", isError: true);
                return;
            }

            var options = new FillBlankOptions();
            if (FillDownRadio.IsChecked == true) options.Direction = FillBlankDirection.Down;
            else if (FillUpRadio.IsChecked == true) options.Direction = FillBlankDirection.Up;
            else if (FillRightRadio.IsChecked == true) options.Direction = FillBlankDirection.Right;
            else if (FillLeftRadio.IsChecked == true) options.Direction = FillBlankDirection.Left;
            else if (FillFixedRadio.IsChecked == true)
            {
                options.Direction = FillBlankDirection.FixedValue;
                options.FixedValue = FixedValueTextBox.Text;
            }

            var result = RangeFill.FillBlanks(matrix, options, out int filledCount);
            _context.SetSelectedValues(result);
            _context.ShowMessage($"Successfully filled {filledCount} blank cells.", "Fill Blank Cells");

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
