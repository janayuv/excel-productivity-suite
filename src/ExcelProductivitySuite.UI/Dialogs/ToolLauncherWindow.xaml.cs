using System;
using System.Collections.Generic;
using System.Linq;
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
        private static readonly List<string> RecentCommandIds = new List<string>();
        private static readonly HashSet<string> FavoriteCommandIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private readonly ICommandRegistry _registry;
        private readonly IExcelContext _context;
        private CommandCategory? _selectedCategory = null;
        private bool _filterRecent = false;
        private bool _filterFavorites = false;

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

            if (_filterRecent)
            {
                commands = commands.Where(c => RecentCommandIds.Contains(c.Id)).ToList();
            }
            else if (_filterFavorites)
            {
                commands = commands.Where(c => c.IsFavorite || FavoriteCommandIds.Contains(c.Id)).ToList();
            }

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

            _filterRecent = false;
            _filterFavorites = false;
            _selectedCategory = null;

            if (CatAll.IsChecked == true) { }
            else if (CatRecent.IsChecked == true) _filterRecent = true;
            else if (CatFavorites.IsChecked == true) _filterFavorites = true;
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
                RecentCommandIds.Remove(cmd.Id);
                RecentCommandIds.Insert(0, cmd.Id);
                if (RecentCommandIds.Count > 10) RecentCommandIds.RemoveAt(10);

                SelectedCommand = cmd;
                DialogResult = true;
                Close();
                cmd.Execute(_context);
            }
        }
    }
}
