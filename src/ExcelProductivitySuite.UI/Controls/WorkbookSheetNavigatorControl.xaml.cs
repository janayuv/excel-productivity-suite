using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.UI.Dialogs;
using ExcelProductivitySuite.UI.ViewModels;

namespace ExcelProductivitySuite.UI.Controls
{
    public partial class WorkbookSheetNavigatorControl : UserControl
    {
        private readonly IExcelContext _context;
        public WorkbookSheetNavigatorViewModel ViewModel { get; }
        public bool IsTaskPaneMode { get; set; }

        public WorkbookSheetNavigatorControl(IExcelContext context, bool isTaskPane = false)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            IsTaskPaneMode = isTaskPane;
            ViewModel = new WorkbookSheetNavigatorViewModel(context);
            DataContext = ViewModel;

            if (isTaskPane)
            {
                PopOutButton.Visibility = Visibility.Visible;
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.Refresh();
        }

        private void ActivateWorkbook_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.ActivateWorkbook();
        }

        private void WorkbooksGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ViewModel.ActivateWorkbook();
        }

        private void NewWorkbook_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.CreateNewWorkbook();
        }

        private void SaveWorkbook_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.SaveWorkbook();
        }

        private void SaveAllWorkbooks_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.SaveAllWorkbooks();
        }

        private void CloseWorkbook_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.CloseWorkbook();
        }

        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.OpenWorkbookFolder();
        }

        private void CopyPath_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.CopyWorkbookPath();
        }

        private void WbSort_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (WbSortCombo == null || ViewModel == null) return;
            string selected = (WbSortCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Open Order";
            ViewModel.WorkbookSort = selected switch
            {
                "Name (A-Z)" => WorkbookSortOrder.NameAscending,
                "Name (Z-A)" => WorkbookSortOrder.NameDescending,
                _ => WorkbookSortOrder.OpenOrder
            };
        }

        private void ActivateWorksheet_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.ActivateWorksheet();
        }

        private void WorksheetsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ViewModel.ActivateWorksheet();
        }

        private void ToggleLast_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.ToggleLastActiveSheet();
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.MoveWorksheet(SheetMoveDirection.Up);
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.MoveWorksheet(SheetMoveDirection.Down);
        }

        private void MoveTop_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.MoveWorksheet(SheetMoveDirection.Top);
        }

        private void MoveEnd_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.MoveWorksheet(SheetMoveDirection.End);
        }

        private void ToggleVisibility_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.ToggleSheetVisibility();
        }

        private void AddSheet_Click(object sender, RoutedEventArgs e)
        {
            var existingNames = ViewModel.Worksheets.Select(s => s.Name).ToList();
            var prompt = new RenameSheetPromptDialog(string.Empty, existingNames, title: "Add New Worksheet")
            {
                Owner = Window.GetWindow(this)
            };

            if (prompt.ShowDialog() == true && !string.IsNullOrWhiteSpace(prompt.ResultName))
            {
                ViewModel.AddNewWorksheet(prompt.ResultName);
            }
        }

        private void RenameSheet_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel.SelectedWorksheet == null) return;
            string oldName = ViewModel.SelectedWorksheet.Name;
            var existingNames = ViewModel.Worksheets.Where(s => !string.Equals(s.Name, oldName, StringComparison.OrdinalIgnoreCase)).Select(s => s.Name).ToList();

            var prompt = new RenameSheetPromptDialog(oldName, existingNames, title: $"Rename '{oldName}'")
            {
                Owner = Window.GetWindow(this)
            };

            if (prompt.ShowDialog() == true && !string.IsNullOrWhiteSpace(prompt.ResultName))
            {
                ViewModel.RenameWorksheet(oldName, prompt.ResultName);
            }
        }

        private void DeleteSheet_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.DeleteWorksheet();
        }

        private void WsFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (WsFilterCombo == null || ViewModel == null) return;
            string selected = (WsFilterCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "All";
            ViewModel.FilterMode = selected switch
            {
                "Visible" => SheetFilterMode.VisibleOnly,
                "Hidden" => SheetFilterMode.HiddenOnly,
                _ => SheetFilterMode.All
            };
        }

        private void PopOut_Click(object sender, RoutedEventArgs e)
        {
            var window = new WorkbookSheetNavigatorWindow(_context);
            window.Show();
        }
    }
}
