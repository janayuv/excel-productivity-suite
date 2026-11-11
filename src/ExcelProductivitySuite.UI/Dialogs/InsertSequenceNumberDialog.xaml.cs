using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;
using Microsoft.Win32;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class SequenceListItemViewModel : INotifyPropertyChanged
    {
        public SequenceNumberDefinition Definition { get; }

        public string Name => Definition.Name;

        public string TypeName => Definition.Type switch
        {
            SequenceType.AlphaUpper => "Letters (A-Z)",
            SequenceType.AlphaLower => "Letters (a-z)",
            SequenceType.RomanUpper => "Roman (I-X)",
            SequenceType.RomanLower => "Roman (i-x)",
            SequenceType.Hexadecimal => "Hex",
            SequenceType.CustomList => "Custom List",
            _ => "Numeric"
        };

        public long StartNumber => Definition.StartNumber;
        public long Increment => Definition.Increment;
        public int Digits => Definition.Digits;
        public bool IsBuiltIn => Definition.IsBuiltIn;

        public string SamplePreview
        {
            get
            {
                var samples = Definition.GeneratePreview(2, DateTime.Now, "Sheet1");
                return samples.Count > 0 ? string.Join(", ", samples) + "..." : "";
            }
        }

        public string NextPreview => Definition.FormatValue(Definition.CurrentNumber, DateTime.Now, "Sheet1");

        public SequenceListItemViewModel(SequenceNumberDefinition definition)
        {
            Definition = definition ?? throw new ArgumentNullException(nameof(definition));
        }

        public void RefreshProperties()
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TypeName)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(StartNumber)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Increment)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Digits)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SamplePreview)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(NextPreview)));
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    public partial class InsertSequenceNumberDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly List<SequenceListItemViewModel> _allSequenceItems = new();
        private readonly ObservableCollection<SequenceListItemViewModel> _sequenceItems = new();
        private readonly ObservableCollection<string> _upcomingPreview = new();
        private bool _isInitialized;

        public InsertSequenceNumberDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            SequenceGrid.ItemsSource = _sequenceItems;
            UpcomingListBox.ItemsSource = _upcomingPreview;

            InitializeControls();
        }

        private void InitializeControls()
        {
            // Fill order combo
            FillOrderCombo.Items.Clear();
            FillOrderCombo.Items.Add(new ComboBoxItem { Content = "Fill vertically down columns (↓)", Tag = SequenceFillOrder.Vertical });
            FillOrderCombo.Items.Add(new ComboBoxItem { Content = "Fill horizontally across rows (→)", Tag = SequenceFillOrder.Horizontal });
            FillOrderCombo.Items.Add(new ComboBoxItem { Content = "Snake / Serpentine vertical (↓ then ↑)", Tag = SequenceFillOrder.SerpentineVertical });
            FillOrderCombo.Items.Add(new ComboBoxItem { Content = "Snake / Serpentine horizontal (→ then ←)", Tag = SequenceFillOrder.SerpentineHorizontal });
            FillOrderCombo.SelectedIndex = 0;

            // Key Column Offset combo for Group Restart
            KeyColumnCombo.Items.Clear();
            KeyColumnCombo.Items.Add(new ComboBoxItem { Content = "Left column (Offset -1)", Tag = -1 });
            KeyColumnCombo.Items.Add(new ComboBoxItem { Content = "Right column (Offset +1)", Tag = 1 });
            KeyColumnCombo.Items.Add(new ComboBoxItem { Content = "2 cols left (Offset -2)", Tag = -2 });
            KeyColumnCombo.Items.Add(new ComboBoxItem { Content = "2 cols right (Offset +2)", Tag = 2 });
            KeyColumnCombo.SelectedIndex = 0;

            // Step Rows combo
            StepRowsCombo.Items.Clear();
            StepRowsCombo.Items.Add(new ComboBoxItem { Content = "Every row (1)", Tag = 1 });
            StepRowsCombo.Items.Add(new ComboBoxItem { Content = "Every 2nd row (2)", Tag = 2 });
            StepRowsCombo.Items.Add(new ComboBoxItem { Content = "Every 3rd row (3)", Tag = 3 });
            StepRowsCombo.Items.Add(new ComboBoxItem { Content = "Every 4th row (4)", Tag = 4 });
            StepRowsCombo.SelectedIndex = 0;

            // Refresh target selection display
            RefreshSelectionState();

            // Load saved sequences
            var loaded = SequenceNumberEngine.LoadSequences();
            _allSequenceItems.Clear();
            foreach (var s in loaded)
            {
                _allSequenceItems.Add(new SequenceListItemViewModel(s));
            }

            _isInitialized = true;
            ApplyFilter();

            if (_sequenceItems.Count > 0)
            {
                SequenceGrid.SelectedIndex = 0;
            }
        }

        private void RefreshSelectionState()
        {
            string sheet = _context.ActiveWorksheetName ?? "ActiveSheet";
            string addr = _context.SelectionAddress ?? "Selection";
            int count = _context.SelectedCellCount;

            SelectionAddressLabel.Text = $"{sheet}!{addr}";
            SelectionCountLabel.Text = count > 0 ? $"({count} cell{(count == 1 ? "" : "s")})" : "";
        }

        private void RefreshSelection_Click(object sender, RoutedEventArgs e)
        {
            RefreshSelectionState();
            UpdateUpcomingPreview();
            StatusLabel.Text = "Selection refreshed.";
        }

        private SequenceListItemViewModel? SelectedViewModel => SequenceGrid.SelectedItem as SequenceListItemViewModel;

        private void SequenceGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdateSelectionState();
        }

        private void SequenceGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (SelectedViewModel != null)
            {
                Edit_Click(sender, e);
            }
        }

        private void FillOrderCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdateUpcomingPreview();
        }

        private void Options_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdateUpcomingPreview();
        }

        private void RestartOnKeyCheck_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            if (KeyColumnCombo != null)
            {
                KeyColumnCombo.IsEnabled = RestartOnKeyCheck.IsChecked == true;
            }
            UpdateUpcomingPreview();
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (!_isInitialized) return;
            string query = SearchBox?.Text?.Trim().ToLowerInvariant() ?? "";
            var selected = SelectedViewModel;

            _sequenceItems.Clear();
            foreach (var vm in _allSequenceItems)
            {
                if (string.IsNullOrEmpty(query) ||
                    vm.Name.ToLowerInvariant().Contains(query) ||
                    vm.TypeName.ToLowerInvariant().Contains(query) ||
                    vm.SamplePreview.ToLowerInvariant().Contains(query))
                {
                    _sequenceItems.Add(vm);
                }
            }

            if (selected != null && _sequenceItems.Contains(selected))
            {
                SequenceGrid.SelectedItem = selected;
            }
            else if (_sequenceItems.Count > 0)
            {
                SequenceGrid.SelectedIndex = 0;
            }

            UpdateSelectionState();
        }

        private void UpdateSelectionState()
        {
            if (!_isInitialized) return;
            var selected = SelectedViewModel;
            if (selected == null)
            {
                EditButton.IsEnabled = false;
                DuplicateButton.IsEnabled = false;
                DeleteButton.IsEnabled = false;
                ResetButton.IsEnabled = false;
                MoveUpButton.IsEnabled = false;
                MoveDownButton.IsEnabled = false;
                _upcomingPreview.Clear();
                return;
            }

            int idx = _allSequenceItems.IndexOf(selected);

            EditButton.IsEnabled = true;
            DuplicateButton.IsEnabled = true;
            DeleteButton.IsEnabled = !selected.IsBuiltIn;
            ResetButton.IsEnabled = true;
            MoveUpButton.IsEnabled = idx > 0;
            MoveDownButton.IsEnabled = idx >= 0 && idx < _allSequenceItems.Count - 1;

            UpdateUpcomingPreview();
        }

        private void UpdateUpcomingPreview()
        {
            if (_upcomingPreview == null) return;
            _upcomingPreview.Clear();

            var selected = SelectedViewModel;
            if (selected == null) return;

            bool continueLast = ContinueFromLastCheck?.IsChecked == true;
            string sheet = _context.ActiveWorksheetName ?? "ActiveSheet";

            var samples = SequenceNumberEngine.GenerateSequence(
                selected.Definition,
                count: 6,
                continueFromCurrent: continueLast,
                referenceDate: DateTime.Now,
                sheetName: sheet,
                out _);

            foreach (var s in samples)
            {
                _upcomingPreview.Add(s);
            }
        }

        private void New_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CreateSequenceNumberDialog();
            dlg.Owner = this;

            if (dlg.ShowDialog() == true)
            {
                var newDef = dlg.ResultSequence;
                var vm = new SequenceListItemViewModel(newDef);
                _allSequenceItems.Add(vm);
                ApplyFilter();
                SequenceGrid.SelectedItem = vm;
                SaveCurrentSequences();
                StatusLabel.Text = $"Created sequence '{newDef.Name}'.";
            }
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            var selected = SelectedViewModel;
            if (selected == null) return;

            var dlg = new CreateSequenceNumberDialog(selected.Definition, isEditMode: true);
            dlg.Owner = this;

            if (dlg.ShowDialog() == true)
            {
                var updated = dlg.ResultSequence;
                selected.Definition.Name = updated.Name;
                selected.Definition.Type = updated.Type;
                selected.Definition.CustomListItems = updated.CustomListItems;
                selected.Definition.Prefix = updated.Prefix;
                selected.Definition.Suffix = updated.Suffix;
                selected.Definition.StartNumber = updated.StartNumber;
                selected.Definition.EndNumber = updated.EndNumber;
                selected.Definition.Increment = updated.Increment;
                selected.Definition.Digits = updated.Digits;
                selected.Definition.RepeatCount = updated.RepeatCount;
                selected.Definition.UseDynamicPlaceholders = updated.UseDynamicPlaceholders;
                selected.Definition.OutputAsNumeric = updated.OutputAsNumeric;

                selected.RefreshProperties();
                UpdateUpcomingPreview();
                SaveCurrentSequences();
                StatusLabel.Text = $"Updated sequence '{updated.Name}'.";
            }
        }

        private void Duplicate_Click(object sender, RoutedEventArgs e)
        {
            var selected = SelectedViewModel;
            if (selected == null) return;

            var copy = selected.Definition.Clone();
            copy.Id = Guid.NewGuid().ToString("N");
            copy.Name = $"{selected.Name} (Copy)";
            copy.IsBuiltIn = false;

            var vm = new SequenceListItemViewModel(copy);
            int idx = _allSequenceItems.IndexOf(selected);
            if (idx >= 0 && idx < _allSequenceItems.Count - 1)
            {
                _allSequenceItems.Insert(idx + 1, vm);
            }
            else
            {
                _allSequenceItems.Add(vm);
            }

            ApplyFilter();
            SequenceGrid.SelectedItem = vm;
            SaveCurrentSequences();
            StatusLabel.Text = $"Duplicated sequence as '{copy.Name}'.";
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            var selected = SelectedViewModel;
            if (selected == null) return;

            if (selected.IsBuiltIn)
            {
                StatusLabel.Text = "Built-in preset sequences cannot be deleted.";
                return;
            }

            var confirm = MessageBox.Show(
                $"Are you sure you want to delete sequence '{selected.Name}'?",
                "Confirm Delete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm == MessageBoxResult.Yes)
            {
                _allSequenceItems.Remove(selected);
                ApplyFilter();
                SaveCurrentSequences();
                StatusLabel.Text = "Sequence deleted.";
            }
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            var selected = SelectedViewModel;
            if (selected == null) return;

            selected.Definition.Reset();
            selected.RefreshProperties();
            UpdateUpcomingPreview();
            SaveCurrentSequences();
            StatusLabel.Text = $"Counter for '{selected.Name}' reset to {selected.Definition.StartNumber}.";
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            var selected = SelectedViewModel;
            if (selected == null) return;

            int idx = _allSequenceItems.IndexOf(selected);
            if (idx > 0)
            {
                _allSequenceItems.RemoveAt(idx);
                _allSequenceItems.Insert(idx - 1, selected);
                ApplyFilter();
                SequenceGrid.SelectedItem = selected;
                SaveCurrentSequences();
            }
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            var selected = SelectedViewModel;
            if (selected == null) return;

            int idx = _allSequenceItems.IndexOf(selected);
            if (idx >= 0 && idx < _allSequenceItems.Count - 1)
            {
                _allSequenceItems.RemoveAt(idx);
                _allSequenceItems.Insert(idx + 1, selected);
                ApplyFilter();
                SequenceGrid.SelectedItem = selected;
                SaveCurrentSequences();
            }
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var openDlg = new OpenFileDialog
            {
                Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                Title = "Import Sequence Definitions"
            };

            if (openDlg.ShowDialog(this) == true)
            {
                if (SequenceNumberEngine.ImportFromFile(openDlg.FileName, out var imported, out string? err))
                {
                    int count = 0;
                    foreach (var s in imported)
                    {
                        var vm = new SequenceListItemViewModel(s);
                        _allSequenceItems.Add(vm);
                        count++;
                    }
                    ApplyFilter();
                    SaveCurrentSequences();
                    StatusLabel.Text = $"Successfully imported {count} sequence definitions.";
                }
                else
                {
                    MessageBox.Show($"Failed to import sequences: {err}", "Import Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            var saveDlg = new SaveFileDialog
            {
                Filter = "JSON Files (*.json)|*.json",
                Title = "Export Sequence Definitions",
                FileName = "CustomSequences.json"
            };

            if (saveDlg.ShowDialog(this) == true)
            {
                var defs = _allSequenceItems.Select(vm => vm.Definition).ToList();
                if (SequenceNumberEngine.ExportToFile(saveDlg.FileName, defs, out string? err))
                {
                    StatusLabel.Text = $"Exported {defs.Count} sequence definitions to {Path.GetFileName(saveDlg.FileName)}.";
                }
                else
                {
                    MessageBox.Show($"Failed to export sequences: {err}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private SequenceFillOrder GetSelectedFillOrder()
        {
            if (FillOrderCombo.SelectedItem is ComboBoxItem item && item.Tag is SequenceFillOrder order)
            {
                return order;
            }
            return SequenceFillOrder.Vertical;
        }

        private int GetSelectedKeyOffset()
        {
            if (KeyColumnCombo.SelectedItem is ComboBoxItem item && item.Tag is int offset)
            {
                return offset;
            }
            return -1;
        }

        private int GetSelectedStepRows()
        {
            if (StepRowsCombo.SelectedItem is ComboBoxItem item && item.Tag is int step)
            {
                return step;
            }
            return 1;
        }

        private bool ExecuteInsertion()
        {
            var selected = SelectedViewModel;
            if (selected == null)
            {
                StatusLabel.Text = "Please select a sequence number definition from the list.";
                return false;
            }

            try
            {
                var options = new SequenceFillOptions
                {
                    FillOrder = GetSelectedFillOrder(),
                    OnlyFillBlankCells = OnlyBlankCheck.IsChecked == true,
                    SkipHiddenCells = SkipHiddenCheck.IsChecked == true,
                    ContinueFromLast = ContinueFromLastCheck.IsChecked == true,
                    RandomizeOrder = RandomizeCheck.IsChecked == true,
                    RestartOnValueChange = RestartOnKeyCheck.IsChecked == true,
                    RestartKeyColumnOffset = GetSelectedKeyOffset(),
                    StepRows = GetSelectedStepRows()
                };

                StatusLabel.Text = "Inserting sequence numbers...";
                var result = _context.InsertSequenceNumbers(selected.Definition, options);

                if (result.Success)
                {
                    selected.RefreshProperties();
                    UpdateUpcomingPreview();
                    SaveCurrentSequences();

                    StatusLabel.Text = result.Message;
                    return true;
                }
                else
                {
                    StatusLabel.Text = $"Operation failed: {result.Message}";
                    return false;
                }
            }
            catch (Exception ex)
            {
                StatusLabel.Text = $"Error: {ex.Message}";
                _context.ShowMessage($"Error inserting sequential numbers: {ex.Message}", "Insert Sequence Error", isError: true);
                return false;
            }
        }

        private void FillRange_Click(object sender, RoutedEventArgs e)
        {
            if (ExecuteInsertion())
            {
                DialogResult = true;
                Close();
            }
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            if (ExecuteInsertion())
            {
                RefreshSelectionState();
            }
        }

        private void SaveCurrentSequences()
        {
            var defs = _allSequenceItems.Select(vm => vm.Definition).ToList();
            SequenceNumberEngine.SaveSequences(defs);
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
