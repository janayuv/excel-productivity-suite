using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Navigation;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ReadingGridFocusDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private int _currentRow = 1;
        private int _currentCol = 1;

        public ReadingGridFocusDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadCurrentContext();
        }

        private void LoadCurrentContext()
        {
            var ctx = _context.GetActiveCellContext();
            _currentRow = ctx.row;
            _currentCol = ctx.col;

            AddressBadge.Text = ctx.address;
            ColLetterText.Text = $"Column {ColumnLetterConverter.IndexToLetter(_currentCol)}";
            ColHeaderText.Text = string.IsNullOrWhiteSpace(ctx.colHeader) ? "(Header Not Found)" : ctx.colHeader;

            RowNumText.Text = $"Row {_currentRow}";
            RowHeaderText.Text = string.IsNullOrWhiteSpace(ctx.rowHeader) ? $"(Row {_currentRow})" : ctx.rowHeader;
        }

        private void NavigateTo(int row, int col)
        {
            if (row < 1) row = 1;
            if (col < 1) col = 1;

            _currentRow = row;
            _currentCol = col;
            _context.SelectRange(_currentRow, _currentCol, 1, 1);
            LoadCurrentContext();
        }

        private void Up_Click(object sender, RoutedEventArgs e) => NavigateTo(_currentRow - 1, _currentCol);
        private void Down_Click(object sender, RoutedEventArgs e) => NavigateTo(_currentRow + 1, _currentCol);
        private void Left_Click(object sender, RoutedEventArgs e) => NavigateTo(_currentRow, _currentCol - 1);
        private void Right_Click(object sender, RoutedEventArgs e) => NavigateTo(_currentRow, _currentCol + 1);

        private void Focus_Click(object sender, RoutedEventArgs e)
        {
            _context.HighlightActiveRowAndColumn(_currentRow, _currentCol);
        }

        private void SelectRow_Click(object sender, RoutedEventArgs e)
        {
            _context.SelectRange(_currentRow, 1, 1, 256);
        }

        private void SelectCol_Click(object sender, RoutedEventArgs e)
        {
            _context.SelectColumn(_currentCol);
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            LoadCurrentContext();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
