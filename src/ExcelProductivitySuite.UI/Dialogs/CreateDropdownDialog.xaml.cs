using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataValidation;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class CreateDropdownDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public CreateDropdownDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            RangeInfoText.Text = $"Target: {_context.SelectionAddress ?? "Active Selection"}";
        }

        private void SourceType_Changed(object sender, RoutedEventArgs e)
        {
            if (InputPromptText == null || ItemsInputBox == null) return;

            if (RangeFormulaRadio?.IsChecked == true)
            {
                InputPromptText.Text = "Enter range reference or named range formula (e.g. =$A$1:$A$10 or =DepartmentList):";
                if (ItemsInputBox.Text.Contains(","))
                {
                    ItemsInputBox.Text = "=$A$1:$A$10";
                }
            }
            else
            {
                InputPromptText.Text = "Enter list items (separated by commas or new lines):";
                if (ItemsInputBox.Text.StartsWith("="))
                {
                    ItemsInputBox.Text = "Approved, Pending, Rejected, Under Review";
                }
            }
        }

        private void ShowAlert_Changed(object sender, RoutedEventArgs e)
        {
            if (AlertFieldsPanel == null) return;
            AlertFieldsPanel.IsEnabled = ShowAlertCheck.IsChecked == true;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before creating dropdown lists.", "Sheet Protected", isError: true);
                return;
            }

            string rawText = ItemsInputBox.Text?.Trim() ?? "";
            bool isFormula = RangeFormulaRadio.IsChecked == true;

            string formattedList;
            if (isFormula)
            {
                formattedList = rawText.StartsWith("=") ? rawText : "=" + rawText;
            }
            else
            {
                formattedList = DropdownListBuilder.BuildListString(rawText, deduplicate: true);
            }

            if (!DropdownListBuilder.ValidateList(formattedList, isFormula, out string? errMsg))
            {
                _context.ShowMessage(errMsg ?? "Invalid dropdown source.", "Create Dropdown List", isError: true);
                return;
            }

            bool showDropdown = ShowDropdownCheck.IsChecked == true;
            bool ignoreBlanks = IgnoreBlanksCheck.IsChecked == true;
            bool showAlert = ShowAlertCheck.IsChecked == true;
            string? alertTitle = showAlert ? AlertTitleBox.Text?.Trim() : null;
            string? alertMsg = showAlert ? AlertMsgBox.Text?.Trim() : null;

            bool success = _context.CreateDropdownList(formattedList, isFormula, ignoreBlanks, showDropdown, alertTitle, alertMsg);
            if (success)
            {
                _context.ShowMessage("In-cell dropdown validation list applied successfully to selection.", "Create Dropdown List", isError: false);
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
