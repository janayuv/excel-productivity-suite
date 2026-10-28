using System;
using System.Collections.Generic;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class UnmergeAndFillDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public class MergedAreaViewModel
        {
            public string Address { get; }
            public string SizeDisplay { get; }
            public string ValueText { get; }

            public MergedAreaViewModel(MergedAreaInfo info)
            {
                Address = info.Address;
                SizeDisplay = $"{info.RowCount}R × {info.ColCount}C ({info.RowCount * info.ColCount} cells)";
                ValueText = string.IsNullOrWhiteSpace(info.ValueText) ? "(Empty / Blank)" : info.ValueText;
            }
        }

        public UnmergeAndFillDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            LoadMergedAreas();
        }

        private void LoadMergedAreas()
        {
            var areas = _context.GetMergedAreasInSelection();
            var viewModels = new List<MergedAreaViewModel>();

            foreach (var a in areas)
            {
                viewModels.Add(new MergedAreaViewModel(a));
            }

            MergedAreasGrid.ItemsSource = viewModels;

            if (viewModels.Count > 0)
            {
                StatusText.Text = $"Found {viewModels.Count} merged block(s) in the current selection. Ready to unmerge and propagate data.";
                UnmergeBtn.IsEnabled = true;
            }
            else
            {
                StatusText.Text = "No merged cells detected in current selection. Select a range containing merged cells in Excel.";
                UnmergeBtn.IsEnabled = false;
            }
        }

        private void UnmergeBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int filled = _context.UnmergeAndFillSelection();
                _context.ShowMessage($"Successfully unmerged selection and filled {filled} constituent cell(s) with primary values.", "Unmerge Complete");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Failed to unmerge and fill: {ex.Message}", "Unmerge Error", isError: true);
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
