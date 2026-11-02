using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ClearValidationDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public ClearValidationDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            TargetInfoText.Text = $"Selection: {_context.SelectionAddress ?? "None"}";
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Cannot modify validation rules.", "Sheet Protected", isError: true);
                return;
            }

            var scope = ValidationClearScope.Selection;
            if (ActiveSheetRadio.IsChecked == true) scope = ValidationClearScope.ActiveWorksheet;
            else if (AllSheetsRadio.IsChecked == true) scope = ValidationClearScope.AllWorksheets;

            int count = _context.ClearValidationRules(scope);
            string scopeDesc = scope switch
            {
                ValidationClearScope.Selection => "selected range",
                ValidationClearScope.ActiveWorksheet => "active worksheet",
                ValidationClearScope.AllWorksheets => $"{count} worksheets in the workbook",
                _ => "target"
            };

            _context.ShowMessage($"Data validation rules cleared successfully from {scopeDesc}.", "Clear Validation", isError: false);
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
