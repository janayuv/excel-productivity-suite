using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Engine.Navigation;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class RenameSheetPromptDialog : BaseDialogWindow
    {
        private readonly IEnumerable<string> _existingNames;
        public string ResultName { get; private set; } = string.Empty;

        public RenameSheetPromptDialog() : this(string.Empty, null, "Rename Worksheet")
        {
        }

        public RenameSheetPromptDialog(string? initialName = null, IEnumerable<string>? existingNames = null, string? title = null)
        {
            InitializeComponent();
            Title = string.IsNullOrWhiteSpace(title) ? "Rename Worksheet" : title!;
            _existingNames = existingNames ?? Array.Empty<string>();
            NameInput.Text = initialName ?? string.Empty;
            NameInput.SelectAll();
            NameInput.Focus();
            ValidateCurrentInput();
        }

        private void NameInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            ValidateCurrentInput();
        }

        private void ValidateCurrentInput()
        {
            if (OkButton == null || ValidationMessage == null) return;

            string candidate = NameInput.Text?.Trim() ?? string.Empty;
            bool valid = SheetNameValidator.Validate(candidate, _existingNames, out string err);

            if (valid)
            {
                ValidationMessage.Text = string.Empty;
                OkButton.IsEnabled = true;
            }
            else
            {
                ValidationMessage.Text = err;
                OkButton.IsEnabled = false;
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            ResultName = NameInput.Text.Trim();
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
