using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Worksheets;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class SequenceSheetsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private List<string> _currentGeneratedNames = new();

        public SequenceSheetsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();
            UpdatePreview();
        }

        private SequenceSheetOptions GetOptions()
        {
            var options = new SequenceSheetOptions();

            if (MonthsRadio?.IsChecked == true) options.Type = SequenceType.Months;
            else if (DaysRadio?.IsChecked == true) options.Type = SequenceType.Days;
            else if (QuartersRadio?.IsChecked == true) options.Type = SequenceType.Quarters;
            else options.Type = SequenceType.Numbers;

            options.Prefix = PrefixBox?.Text ?? "";
            options.Suffix = SuffixBox?.Text ?? "";

            int start = 1;
            if (int.TryParse(StartBox?.Text, out int s)) start = Math.Max(1, s);
            options.StartNumber = start;

            int count = 5;
            if (int.TryParse(CountBox?.Text, out int c)) count = Math.Max(1, Math.Min(c, 100));
            options.Count = count;

            options.ZeroPadding = PadZerosCheck?.IsChecked == true ? 2 : 0;
            options.FullName = FullNameCheck?.IsChecked == true;

            return options;
        }

        private void UpdatePreview()
        {
            if (PreviewListBox == null || StatusText == null || CreateButton == null) return;

            var options = GetOptions();
            var names = SequenceSheetGenerator.GenerateNames(options);
            _currentGeneratedNames = names;

            var existing = _context.GetWorksheetNames();
            var (valid, errors) = SequenceSheetGenerator.ValidateGeneratedNames(names, existing);

            PreviewListBox.ItemsSource = names;
            if (PreviewCountText != null)
            {
                PreviewCountText.Text = $"{names.Count} sheet(s) proposed ({valid.Count} valid)";
            }

            if (errors.Count > 0)
            {
                StatusText.Text = $"Warning: {errors[0]}";
                CreateButton.IsEnabled = valid.Count > 0;
            }
            else
            {
                StatusText.Text = $"All {names.Count} sheet name(s) are valid and ready to create.";
                CreateButton.IsEnabled = true;
            }
        }

        private void Option_Changed(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void Input_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsWorkbookProtected)
            {
                _context.ShowMessage("The active workbook structure is protected. Please unprotect the workbook before inserting new worksheets.", "Workbook Protected", isError: true);
                return;
            }

            var existing = _context.GetWorksheetNames();
            var (valid, errors) = SequenceSheetGenerator.ValidateGeneratedNames(_currentGeneratedNames, existing);

            if (valid.Count == 0)
            {
                _context.ShowMessage("None of the proposed sheet names are valid or all already exist.", "Invalid Sheet Names", isError: true);
                return;
            }

            int created = 0;
            var createdErrors = new List<string>();

            foreach (var sheetName in valid)
            {
                try
                {
                    bool ok = _context.CreateWorksheetWithData(sheetName, new object?[0, 0], activate: false);
                    if (ok) created++;
                }
                catch (Exception ex)
                {
                    createdErrors.Add($"Could not create '{sheetName}': {ex.Message}");
                }
            }

            if (created > 0)
            {
                _context.ShowMessage($"Successfully created {created} sequentially named worksheet(s).", "Sheets Created");
                Close();
            }
            else
            {
                string msg = createdErrors.Count > 0 ? string.Join("\n", createdErrors) : "Failed to create worksheets.";
                _context.ShowMessage(msg, "Error", isError: true);
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
