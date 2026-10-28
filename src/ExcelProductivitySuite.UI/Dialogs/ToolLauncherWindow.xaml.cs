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
        private bool _filterActiveOnly = false;
        private bool _filterRoadmapOnly = false;
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
            if (CommandsListBox == null) return;

            string query = SearchTextBox?.Text ?? string.Empty;
            FeatureStatus? statusFilter = null;
            if (_filterActiveOnly) statusFilter = FeatureStatus.ImplementedAndVerified;
            else if (_filterRoadmapOnly) statusFilter = FeatureStatus.Planned;

            var commands = _registry.SearchCommands(query, _selectedCategory, statusFilter);

            if (_filterRecent)
            {
                commands = commands.Where(c => RecentCommandIds.Contains(c.Id)).ToList();
            }
            else if (_filterFavorites)
            {
                commands = commands.Where(c => c.IsFavorite || FavoriteCommandIds.Contains(c.Id)).ToList();
            }

            CommandsListBox.ItemsSource = commands;

            int activeCount = commands.Count(c => c.IsImplemented);
            int roadmapCount = commands.Count - activeCount;

            if (StatusCountText != null)
            {
                StatusCountText.Text = $"{commands.Count} tools displayed ({activeCount} active, {roadmapCount} in roadmap)";
            }

            if (commands.Count > 0)
            {
                CommandsListBox.SelectedIndex = 0;
            }

            if (ClearButton != null)
            {
                ClearButton.Visibility = string.IsNullOrEmpty(query) ? Visibility.Collapsed : Visibility.Visible;
            }
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

        private void Filter_Checked(object sender, RoutedEventArgs e)
        {
            _filterActiveOnly = false;
            _filterRoadmapOnly = false;
            _filterRecent = false;
            _filterFavorites = false;

            if (FilterAll?.IsChecked == true) { }
            else if (FilterActive?.IsChecked == true) _filterActiveOnly = true;
            else if (FilterRoadmap?.IsChecked == true) _filterRoadmapOnly = true;
            else if (FilterRecent?.IsChecked == true) _filterRecent = true;
            else if (FilterFavorites?.IsChecked == true) _filterFavorites = true;

            RefreshList();
        }

        private void CategoryComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CategoryComboBox?.SelectedItem is ComboBoxItem item)
            {
                if (item.Tag is string tag && Enum.TryParse<CommandCategory>(tag, out var cat))
                {
                    _selectedCategory = cat;
                }
                else
                {
                    _selectedCategory = null;
                }
            }
            else
            {
                _selectedCategory = null;
            }

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
                if (RecentCommandIds.Count > 15) RecentCommandIds.RemoveAt(15);

                SelectedCommand = cmd;

                if (cmd.IsImplemented)
                {
                    DialogResult = true;
                    Close();
                    cmd.Execute(_context);
                }
                else
                {
                    // For roadmap commands, display roadmap information modal without closing launcher
                    cmd.Execute(_context);
                }
            }
        }
    }
}
