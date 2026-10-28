using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class MergeTablesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly object?[,]? _primaryMatrix;
        private object?[,]? _lookupMatrix;

        public MergeTablesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            _primaryMatrix = _context.GetSelectedValues();
            PrimaryAddressLabel.Text = _context.SelectionAddress ?? "Selection";

            LoadSheets();
            PopulatePrimaryColumns();
        }

        private void LoadSheets()
        {
            LookupSheetCombo.Items.Clear();
            var sheets = _context.GetWorksheetNames();
            foreach (var s in sheets)
            {
                LookupSheetCombo.Items.Add(s);
            }

            if (LookupSheetCombo.Items.Count > 1)
            {
                // Select 2nd sheet by default if available
                LookupSheetCombo.SelectedIndex = 1;
            }
            else if (LookupSheetCombo.Items.Count > 0)
            {
                LookupSheetCombo.SelectedIndex = 0;
            }
        }

        private void PopulatePrimaryColumns()
        {
            PrimaryKeyCombo.Items.Clear();
            if (_primaryMatrix == null) return;

            int rStart = _primaryMatrix.GetLowerBound(0);
            int cStart = _primaryMatrix.GetLowerBound(1);
            int cEnd = _primaryMatrix.GetUpperBound(1);
            int cols = cEnd - cStart + 1;
            bool hasHeaders = HasHeadersCheck?.IsChecked == true;

            for (int c = 0; c < cols; c++)
            {
                string name = hasHeaders ? _primaryMatrix[rStart, cStart + c]?.ToString() ?? "" : "";
                if (string.IsNullOrWhiteSpace(name)) name = $"Column {c + 1}";
                PrimaryKeyCombo.Items.Add(name);
            }

            if (PrimaryKeyCombo.Items.Count > 0) PrimaryKeyCombo.SelectedIndex = 0;
        }

        private void LookupSheet_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            LookupKeyCombo.Items.Clear();
            if (LookupSheetCombo.SelectedItem is string sheetName)
            {
                _lookupMatrix = _context.GetActiveWorksheetRangeValues(sheetName);
                if (_lookupMatrix != null)
                {
                    int rStart = _lookupMatrix.GetLowerBound(0);
                    int cStart = _lookupMatrix.GetLowerBound(1);
                    int cEnd = _lookupMatrix.GetUpperBound(1);
                    int cols = cEnd - cStart + 1;
                    bool hasHeaders = HasHeadersCheck?.IsChecked == true;

                    for (int c = 0; c < cols; c++)
                    {
                        string name = hasHeaders ? _lookupMatrix[rStart, cStart + c]?.ToString() ?? "" : "";
                        if (string.IsNullOrWhiteSpace(name)) name = $"Column {c + 1}";
                        LookupKeyCombo.Items.Add(name);
                    }

                    if (LookupKeyCombo.Items.Count > 0) LookupKeyCombo.SelectedIndex = 0;
                }
            }
        }

        private void HasHeaders_Changed(object sender, RoutedEventArgs e)
        {
            PopulatePrimaryColumns();
            LookupSheet_Changed(sender, null!);
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Merge_Click(object sender, RoutedEventArgs e)
        {
            if (_primaryMatrix == null || _lookupMatrix == null)
            {
                _context.ShowMessage("Primary and secondary tables must both contain data.", "Missing Table", isError: true);
                return;
            }

            var options = new TableMergeOptions
            {
                PrimaryKeyColumn = Math.Max(0, PrimaryKeyCombo.SelectedIndex),
                LookupKeyColumn = Math.Max(0, LookupKeyCombo.SelectedIndex),
                HasHeaders = HasHeadersCheck.IsChecked == true,
                AppendUnmatchedRows = AppendUnmatchedCheck.IsChecked == true,
                MatchCase = MatchCaseCheck.IsChecked == true
            };

            var merged = TableMerger.Merge(_primaryMatrix, _lookupMatrix, options, out int matched, out int unmatched);

            string sheetName = "Merged_Table";
            bool success = _context.CreateWorksheetWithData(sheetName, merged, activate: true);

            if (success)
            {
                _context.ShowMessage($"Successfully merged tables into '{sheetName}'.\nMatched rows: {matched}\nUnmatched primary rows: {unmatched}", "Merge Complete");
                DialogResult = true;
                Close();
            }
        }
    }
}
