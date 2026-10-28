using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class CellContentInspectorDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public CellContentInspectorDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadCellData();
        }

        private void LoadCellData()
        {
            var details = _context.GetActiveCellDetails();
            if (details == null)
            {
                AddressBadge.Text = _context.SelectionAddress ?? "$A$1";
                SheetBadge.Text = _context.ActiveWorksheetName ?? "Sheet1";
                TypeBadge.Text = "Empty";
                CharCountText.Text = "Chars: 0";
                WordCountText.Text = "Words: 0";
                LineCountText.Text = "Lines: 0";
                ContentBox.Text = "(No cell selected or cell is empty)";
                FormulaGroup.Visibility = Visibility.Collapsed;
                CopyFormulaBtn.IsEnabled = false;
                return;
            }

            AddressBadge.Text = details.Address;
            SheetBadge.Text = details.SheetName;
            TypeBadge.Text = details.DataType;

            CharCountText.Text = $"Chars: {details.CharCount:N0}";
            WordCountText.Text = $"Words: {details.WordCount:N0}";
            LineCountText.Text = $"Lines: {details.LineCount:N0}";

            ContentBox.Text = details.ValueText;

            if (!string.IsNullOrWhiteSpace(details.FormulaText))
            {
                FormulaGroup.Visibility = Visibility.Visible;
                FormulaBox.Text = details.FormulaText;
                CopyFormulaBtn.IsEnabled = true;
            }
            else
            {
                FormulaGroup.Visibility = Visibility.Collapsed;
                CopyFormulaBtn.IsEnabled = false;
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            LoadCellData();
        }

        private void CopyContent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Clipboard.SetText(ContentBox.Text);
                _context.ShowMessage("Cell content copied to clipboard.", "Clipboard");
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Could not copy: {ex.Message}", "Clipboard Error", isError: true);
            }
        }

        private void CopyFormula_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(FormulaBox.Text))
                {
                    Clipboard.SetText(FormulaBox.Text);
                    _context.ShowMessage("Formula expression copied to clipboard.", "Clipboard");
                }
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Could not copy: {ex.Message}", "Clipboard Error", isError: true);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
