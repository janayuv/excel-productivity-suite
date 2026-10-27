using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.UI.Controls;
using ExcelProductivitySuite.UI.Models;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class SplitTextDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<string> _sampleOriginals;

        public SplitTextDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _sampleOriginals = _context.GetSampleOriginalValues(10);
            UpdatePreview();
        }

        private TextSplitOptions GetOptions()
        {
            var options = new TextSplitOptions
            {
                TrimTokens = TrimTokensCheckBox.IsChecked == true,
                RemoveEmptyTokens = RemoveEmptyCheckBox.IsChecked == true
            };

            if (DelimComma.IsChecked == true) options.Delimiter = ",";
            else if (DelimSpace.IsChecked == true) options.Delimiter = " ";
            else if (DelimSemicolon.IsChecked == true) options.Delimiter = ";";
            else if (DelimNewline.IsChecked == true) options.SplitByNewline = true;
            else if (DelimCustom.IsChecked == true) options.Delimiter = CustomDelimTextBox.Text;

            return options;
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            if (CustomDelimTextBox != null)
            {
                CustomDelimTextBox.IsEnabled = DelimCustom.IsChecked == true;
            }

            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewTable == null) return;

            var options = GetOptions();
            var previews = new List<PreviewItem>();

            for (int i = 0; i < _sampleOriginals.Count; i++)
            {
                string original = _sampleOriginals[i];
                string[] tokens = TextSplitter.Split(original, options);
                string preview = string.Join(" │ ", tokens);

                previews.Add(new PreviewItem
                {
                    RowIndex = i + 1,
                    Original = original,
                    Result = preview
                });
            }

            PreviewTable.ItemsSource = previews;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Split Cells", isError: true);
                return;
            }

            var options = GetOptions();
            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rows = rEnd - rStart + 1;
            int cols = cEnd - cStart + 1;

            if (SplitToColumnsRadio.IsChecked == true)
            {
                // Find maximum tokens per row
                var rowTokens = new List<string[]>();
                int maxTokens = 1;

                for (int r = rStart; r <= rEnd; r++)
                {
                    string str = matrix[r, cStart]?.ToString() ?? "";
                    string[] tokens = TextSplitter.Split(str, options);
                    if (tokens.Length == 0) tokens = new[] { "" };
                    rowTokens.Add(tokens);
                    if (tokens.Length > maxTokens) maxTokens = tokens.Length;
                }

                object?[,] output = new object?[rows, maxTokens];
                for (int i = 0; i < rows; i++)
                {
                    string[] tokens = rowTokens[i];
                    for (int t = 0; t < maxTokens; t++)
                    {
                        output[i, t] = t < tokens.Length ? tokens[t] : "";
                    }
                }

                _context.SetRangeValues(rStart, cStart, output);
                _context.ShowMessage($"Split data into {maxTokens} columns successfully.", "Split Cells");
            }
            else // Split to rows
            {
                var allRows = new List<string>();
                for (int r = rStart; r <= rEnd; r++)
                {
                    string str = matrix[r, cStart]?.ToString() ?? "";
                    string[] tokens = TextSplitter.Split(str, options);
                    if (tokens.Length == 0) allRows.Add("");
                    else allRows.AddRange(tokens);
                }

                object?[,] output = new object?[allRows.Count, 1];
                for (int i = 0; i < allRows.Count; i++)
                {
                    output[i, 0] = allRows[i];
                }

                _context.SetRangeValues(rStart, cStart, output);
                _context.ShowMessage($"Split data into {allRows.Count} rows successfully.", "Split Cells");
            }

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
