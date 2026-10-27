using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ToolLauncherWindow : BaseDialogWindow
    {
        private readonly ICommandRegistry _registry;
        private readonly IExcelContext _context;
        private CommandCategory? _selectedCategory = null;

        public SuiteCommand? SelectedCommand { get; private set; }

        public ToolLauncherWindow(ICommandRegistry registry, IExcelContext context)
        {
            _registry = registry ?? throw new ArgumentNullException(nameof(registry));
            _context = context ?? throw new ArgumentNullException(nameof(context));

            InitializeComponent();
            RefreshList();
        }

        private void RefreshList()
        {
            string query = SearchTextBox.Text;
            var commands = _registry.SearchCommands(query, _selectedCategory);

            CommandsListBox.ItemsSource = commands;
            StatusCountText.Text = $"{commands.Count} tools available";

            if (commands.Count > 0)
            {
                CommandsListBox.SelectedIndex = 0;
            }

            ClearButton.Visibility = string.IsNullOrEmpty(query) ? Visibility.Collapsed : Visibility.Visible;
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshList();
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            SearchTextBox.Text = "";
            SearchTextBox.Focus();
        }

        private void CategoryFilter_Checked(object sender, RoutedEventArgs e)
        {
            if (CatAll == null) return;

            if (CatAll.IsChecked == true) _selectedCategory = null;
            else if (CatText.IsChecked == true) _selectedCategory = CommandCategory.Text;
            else if (CatRange.IsChecked == true) _selectedCategory = CommandCategory.RangeAndCells;
            else if (CatCleaning.IsChecked == true) _selectedCategory = CommandCategory.DataCleaning;
            else if (CatSheets.IsChecked == true) _selectedCategory = CommandCategory.Worksheet;
            else if (CatFormulas.IsChecked == true) _selectedCategory = CommandCategory.Formulas;

            RefreshList();
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteSelectedCommand();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void CommandsListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ExecuteSelectedCommand();
        }

        private void CommandsListBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ExecuteSelectedCommand();
                e.Handled = true;
            }
        }

        private void SearchTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down && CommandsListBox.Items.Count > 0)
            {
                CommandsListBox.Focus();
                if (CommandsListBox.SelectedIndex < 0) CommandsListBox.SelectedIndex = 0;
                e.Handled = true;
            }
        }

        private void ExecuteSelectedCommand()
        {
            if (CommandsListBox.SelectedItem is SuiteCommand cmd)
            {
                SelectedCommand = cmd;
                DialogResult = true;
                Close();
                cmd.Execute(_context);
            }
        }
    }
}
