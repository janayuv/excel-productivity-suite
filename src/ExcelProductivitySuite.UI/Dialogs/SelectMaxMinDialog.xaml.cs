using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class SelectMaxMinDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _matrix;
        private List<(int Row, int Column, double Value)>? _extremeCells;

        public SelectMaxMinDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _matrix = _context.GetSelectedValues();
            UpdateSummary();
        }

        private MaxMinSelectionOptions GetOptions()
        {
            var options = new MaxMinSelectionOptions
            {
                Target = MaxRadio?.IsChecked == true ? MaxMinTarget.Maximum : MaxMinTarget.Minimum
            };

            if (EachRowRadio?.IsChecked == true)
            {
                options.Scope = MaxMinScope.EachRow;
            }
            else if (EachColRadio?.IsChecked == true)
            {
                options.Scope = MaxMinScope.EachColumn;
            }
            else
            {
                options.Scope = MaxMinScope.EntireSelection;
            }

            if (int.TryParse(CountBox?.Text?.Trim(), out int n) && n > 0)
            {
                options.Count = n;
            }

            return options;
        }

        private void UpdateSummary()
        {
            if (SummaryText == null) return;
            if (_matrix == null)
            {
                SummaryText.Text = "No range selected.";
                return;
            }

            try
            {
                var options = GetOptions();
                int startRow = _context.SelectionStartRow;
                int startCol = _context.SelectionStartColumn;

                _extremeCells = MaxMinSelector.FindExtremeCells(_matrix, startRow, startCol, options);

                string targetDesc = options.Target == MaxMinTarget.Maximum ? "maximum" : "minimum";
                SummaryText.Text = $"Found {_extremeCells.Count} {targetDesc} cell(s) ready to select.";
            }
            catch (Exception ex)
            {
                SummaryText.Text = "Error: " + ex.Message;
            }
        }

        private void OptionChanged(object sender, RoutedEventArgs e)
        {
            UpdateSummary();
        }

        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            if (_extremeCells == null || _extremeCells.Count == 0)
            {
                _context.ShowMessage("No numeric cells found matching criteria.", "Select Max/Min");
                return;
            }

            var coords = _extremeCells.Select(c => (c.Row, c.Column)).ToList();
            _context.SelectCells(coords);

            DialogResult = true;
            Close();
            _context.ShowMessage($"Selected {coords.Count} extreme cell(s) in Excel.", "Select Max/Min");
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
