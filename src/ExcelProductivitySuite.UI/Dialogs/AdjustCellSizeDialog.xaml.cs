using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formatting;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class AdjustCellSizeDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private CellDimensionUnit _currentUnit = CellDimensionUnit.Points;

        public AdjustCellSizeDialog(IExcelContext context, bool focusRowHeight = false, bool focusColWidth = false)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SelectionInfoText.Text = $"Target: {_context.SelectionAddress ?? "Selection"}";

            if (focusRowHeight && !focusColWidth)
            {
                RowHeightCheck.IsChecked = true;
                ColWidthCheck.IsChecked = false;
            }
            else if (focusColWidth && !focusRowHeight)
            {
                RowHeightCheck.IsChecked = false;
                ColWidthCheck.IsChecked = true;
            }
            else
            {
                RowHeightCheck.IsChecked = true;
                ColWidthCheck.IsChecked = true;
            }
        }

        private CellDimensionUnit GetSelectedUnit()
        {
            if (UnitInchesRadio?.IsChecked == true) return CellDimensionUnit.Inches;
            if (UnitCmRadio?.IsChecked == true) return CellDimensionUnit.Centimeters;
            return CellDimensionUnit.Points;
        }

        private void Unit_Changed(object sender, RoutedEventArgs e)
        {
            if (RowHeightBox == null || ColWidthBox == null) return;

            var newUnit = GetSelectedUnit();
            if (newUnit == _currentUnit) return;

            // Convert existing displayed values
            if (double.TryParse(RowHeightBox.Text, out double rh))
            {
                double points = CellDimensionHelper.ToPoints(rh, _currentUnit);
                double converted = CellDimensionHelper.FromPoints(points, newUnit);
                RowHeightBox.Text = Math.Round(converted, 2).ToString();
            }

            if (double.TryParse(ColWidthBox.Text, out double cw))
            {
                double points = CellDimensionHelper.ToPoints(cw, _currentUnit);
                double converted = CellDimensionHelper.FromPoints(points, newUnit);
                ColWidthBox.Text = Math.Round(converted, 2).ToString();
            }

            _currentUnit = newUnit;
        }

        private void RowHeightCheck_Changed(object sender, RoutedEventArgs e)
        {
            if (RowHeightBox == null) return;
            RowHeightBox.IsEnabled = RowHeightCheck.IsChecked == true && AutoFitRowsCheck.IsChecked != true;
        }

        private void ColWidthCheck_Changed(object sender, RoutedEventArgs e)
        {
            if (ColWidthBox == null) return;
            ColWidthBox.IsEnabled = ColWidthCheck.IsChecked == true && AutoFitColsCheck.IsChecked != true;
        }

        private void AutoFitRows_Changed(object sender, RoutedEventArgs e)
        {
            if (RowHeightBox == null) return;
            if (AutoFitRowsCheck.IsChecked == true)
            {
                RowHeightCheck.IsChecked = true;
                RowHeightBox.IsEnabled = false;
            }
            else
            {
                RowHeightBox.IsEnabled = RowHeightCheck.IsChecked == true;
            }
        }

        private void AutoFitCols_Changed(object sender, RoutedEventArgs e)
        {
            if (ColWidthBox == null) return;
            if (AutoFitColsCheck.IsChecked == true)
            {
                ColWidthCheck.IsChecked = true;
                ColWidthBox.IsEnabled = false;
            }
            else
            {
                ColWidthBox.IsEnabled = ColWidthCheck.IsChecked == true;
            }
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_context.IsCurrentSheetProtected)
            {
                _context.ShowMessage("The active worksheet is protected. Please unprotect the sheet before adjusting cell dimensions.", "Sheet Protected", isError: true);
                return;
            }

            bool adjustRow = RowHeightCheck.IsChecked == true;
            bool adjustCol = ColWidthCheck.IsChecked == true;
            bool autoFitRows = AutoFitRowsCheck.IsChecked == true;
            bool autoFitCols = AutoFitColsCheck.IsChecked == true;
            bool entireSheet = EntireSheetScopeRadio.IsChecked == true;

            if (!adjustRow && !adjustCol)
            {
                _context.ShowMessage("Please select either Row Height or Column Width to adjust.", "Adjust Cell Size", isError: true);
                return;
            }

            double? rowHeightPoints = null;
            if (adjustRow && !autoFitRows)
            {
                if (!double.TryParse(RowHeightBox.Text.Trim(), out double rhVal) || rhVal <= 0)
                {
                    _context.ShowMessage("Please enter a valid positive number for Row Height.", "Adjust Cell Size", isError: true);
                    return;
                }
                rowHeightPoints = CellDimensionHelper.ClampRowHeight(CellDimensionHelper.ToPoints(rhVal, _currentUnit));
            }

            double? colWidthPoints = null;
            if (adjustCol && !autoFitCols)
            {
                if (!double.TryParse(ColWidthBox.Text.Trim(), out double cwVal) || cwVal <= 0)
                {
                    _context.ShowMessage("Please enter a valid positive number for Column Width.", "Adjust Cell Size", isError: true);
                    return;
                }
                colWidthPoints = CellDimensionHelper.ClampColumnWidth(CellDimensionHelper.ToPoints(cwVal, _currentUnit));
            }

            bool success = _context.AdjustCellDimensions(rowHeightPoints, colWidthPoints, autoFitRows, autoFitCols, entireSheet);
            if (success)
            {
                _context.ShowMessage("Cell dimensions adjusted successfully.", "Adjust Cell Size", isError: false);
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
