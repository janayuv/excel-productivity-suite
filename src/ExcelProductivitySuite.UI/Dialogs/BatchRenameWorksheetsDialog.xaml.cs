using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Worksheets;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class BatchRenameWorksheetsDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly IReadOnlyList<string> _originalNames;

        public BatchRenameWorksheetsDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _originalNames = _context.GetWorksheetNames();
            UpdatePreview();
        }

        private BatchRenameOptions GetOptions()
        {
            return new BatchRenameOptions
            {
                Prefix = PrefixTextBox.Text,
                Suffix = SuffixTextBox.Text,
                FindText = FindTextBox.Text,
                ReplaceWith = ReplaceTextBox.Text,
                AppendSequenceNumber = SequenceCheckBox.IsChecked == true,
                SequenceStart = 1,
                SequenceFormat = "D2"
            };
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void OptionChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewGrid == null || _originalNames == null) return;

            var options = GetOptions();
            var renamedMap = WorksheetNameValidator.GenerateRenamedNames(_originalNames, options);

            var previews = new List<PreviewItem>();
            int idx = 1;
            foreach (var orig in _originalNames)
            {
                previews.Add(new PreviewItem
                {
                    RowIndex = idx++,
                    Original = orig,
                    Result = renamedMap.TryGetValue(orig, out var renamed) ? renamed : orig
                });
            }

            PreviewGrid.ItemsSource = previews;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            var options = GetOptions();
            var renameMap = WorksheetNameValidator.GenerateRenamedNames(_originalNames, options);

            _context.RenameWorksheets(renameMap);
            _context.ShowMessage($"Successfully renamed {renameMap.Count} worksheets.", "Batch Rename Worksheets");

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
