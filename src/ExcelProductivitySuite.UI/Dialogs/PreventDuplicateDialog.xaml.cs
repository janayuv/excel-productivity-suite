using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class PreventDuplicateDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public PreventDuplicateDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            LoadTargetInfo();
        }

        private void LoadTargetInfo()
        {
            string addr = _context.SelectionAddress ?? "Selection";
            RangeAddressLabel.Text = addr;

            int col = _context.SelectionStartColumn;
            int row = _context.SelectionStartRow;
            string colLetter = DuplicateValidationHelper.IndexToColumnLetter(col);
            string formula = DuplicateValidationHelper.BuildCountIfFormula(colLetter, row);

            FormulaPreviewLabel.Text = formula;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            string errorMsg = ErrorMessageBox.Text.Trim();
            bool success = _context.ApplyPreventDuplicateValidation(errorMsg);

            if (success)
            {
                _context.ShowMessage("Duplicate prevention validation rule successfully applied to selection.", "Validation Active");
                DialogResult = true;
                Close();
            }
        }
    }
}
