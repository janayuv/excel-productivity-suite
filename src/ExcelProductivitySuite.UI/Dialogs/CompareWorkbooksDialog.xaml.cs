using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Worksheets;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class WorkbookDiffDisplayItem
    {
        public string SheetName { get; set; } = "";
        public string CellAddress { get; set; } = "";
        public string ValueA { get; set; } = "";
        public string ValueB { get; set; } = "";
        public string DiffType { get; set; } = "";
    }

    public partial class CompareWorkbooksDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private WorkbookComparisonResult? _lastResult;

        public CompareWorkbooksDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            PopulateWorkbooks();
        }

        private void PopulateWorkbooks()
        {
            var books = _context.GetOpenWorkbookNames();
            WorkbookAComboBox.Items.Clear();
            WorkbookBComboBox.Items.Clear();

            foreach (var b in books)
            {
                WorkbookAComboBox.Items.Add(b);
                WorkbookBComboBox.Items.Add(b);
            }

            if (WorkbookAComboBox.Items.Count > 0) WorkbookAComboBox.SelectedIndex = 0;
            if (WorkbookBComboBox.Items.Count > 1) WorkbookBComboBox.SelectedIndex = 1;
            else if (WorkbookBComboBox.Items.Count > 0) WorkbookBComboBox.SelectedIndex = 0;
        }

        private void CompareButton_Click(object sender, RoutedEventArgs e)
        {
            string? wbA = WorkbookAComboBox.SelectedItem as string;
            string? wbB = WorkbookBComboBox.SelectedItem as string;

            if (string.IsNullOrEmpty(wbA) || string.IsNullOrEmpty(wbB))
            {
                _context.ShowMessage("Please select two workbooks to compare.", "Compare Workbooks", isError: true);
                return;
            }

            // Read sheets from Workbook A
            var sheetsA = _context.GetWorkbookSheetNames(wbA!);
            var aMatrices = new Dictionary<string, object?[,]>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in sheetsA)
            {
                var m = _context.GetWorkbookRangeValues(wbA!, s);
                if (m != null) aMatrices[s] = m;
            }

            // Read sheets from Workbook B
            var sheetsB = _context.GetWorkbookSheetNames(wbB!);
            var bMatrices = new Dictionary<string, object?[,]>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in sheetsB)
            {
                var m = _context.GetWorkbookRangeValues(wbB!, s);
                if (m != null) bMatrices[s] = m;
            }

            var options = new WorkbookCompareOptions
            {
                MatchCase = MatchCaseCheckBox.IsChecked == true,
                TrimWhitespace = TrimSpacesCheckBox.IsChecked == true,
                IgnoreEmptyCells = IgnoreEmptyCheckBox.IsChecked == true
            };

            _lastResult = WorkbookComparer.Compare(wbA!, aMatrices, wbB!, bMatrices, options);

            var items = new List<WorkbookDiffDisplayItem>();
            foreach (var report in _lastResult.SheetReports)
            {
                foreach (var diff in report.Discrepancies)
                {
                    items.Add(new WorkbookDiffDisplayItem
                    {
                        SheetName = report.SheetName,
                        CellAddress = $"R{diff.Row}C{diff.Column}",
                        ValueA = diff.SourceValue ?? "(Empty)",
                        ValueB = diff.TargetValue ?? "(Empty)",
                        DiffType = diff.DifferenceType.ToString()
                    });
                }
            }

            FindingsDataGrid.ItemsSource = items;

            if (ExportReportButton != null) ExportReportButton.IsEnabled = true;

            int missingSheets = _lastResult.SheetsOnlyInA.Count + _lastResult.SheetsOnlyInB.Count;
            SummaryText.Text = $"Compared {sheetsA.Count} vs {sheetsB.Count} sheets. Found {_lastResult.TotalDiscrepancies} cell discrepancies across {missingSheets} missing sheet(s).";
        }

        private void ExportReportButton_Click(object sender, RoutedEventArgs e)
        {
            if (_lastResult == null) return;

            if (_context.IsWorkbookProtected)
            {
                _context.ShowMessage("The active workbook structure is protected. Cannot create report worksheet.", "Workbook Protected", isError: true);
                return;
            }

            var reportMatrix = _lastResult.BuildSummaryMatrix();
            string sheetName = "Workbook_Diff_Report";
            bool ok = _context.CreateWorksheetWithData(sheetName, reportMatrix);

            if (ok)
            {
                _context.ShowMessage($"Exported workbook comparison report to worksheet '{sheetName}'.", "Report Generated");
                Close();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
