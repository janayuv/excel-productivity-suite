using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class SequenceListItemViewModel : INotifyPropertyChanged
    {
        public SequenceNumberDefinition Definition { get; }

        public string Name => Definition.Name;
        public long StartNumber => Definition.StartNumber;
        public long Increment => Definition.Increment;
        public int Digits => Definition.Digits;
        public bool IsBuiltIn => Definition.IsBuiltIn;

        public string SamplePreview
        {
            get
            {
                var samples = Definition.GeneratePreview(2);
                return samples.Count > 0 ? string.Join(", ", samples) + "..." : "";
            }
        }

        public string NextPreview => Definition.FormatValue(Definition.CurrentNumber);

        public SequenceListItemViewModel(SequenceNumberDefinition definition)
        {
            Definition = definition ?? throw new ArgumentNullException(nameof(definition));
        }

        public void RefreshProperties()
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Name)));
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
        private readonly ObservableCollection<SequenceListItemViewModel> _sequenceItems = new();
        private readonly ObservableCollection<string> _upcomingPreview = new();

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
            FillOrderCombo.Items.Add("Fill vertically cell after cell");
            FillOrderCombo.Items.Add("Fill horizontally cell after cell");
            FillOrderCombo.SelectedIndex = 0;

            // Target Selection display
            string sheet = _context.ActiveWorksheetName ?? "ActiveSheet";
            string addr = _context.SelectionAddress ?? "Selection";
            SelectionAddressLabel.Text = $"{sheet}!{addr}";

            // Load saved sequences
            var loaded = SequenceNumberEngine.LoadSequences();
            _sequenceItems.Clear();
            foreach (var s in loaded)
            {
                _sequenceItems.Add(new SequenceListItemViewModel(s));
            }

            if (_sequenceItems.Count > 0)
            {
                SequenceGrid.SelectedIndex = 0;
            }
        }

        private SequenceListItemViewModel? SelectedViewModel => SequenceGrid.SelectedItem as SequenceListItemViewModel;

        private void SequenceGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
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
            UpdateUpcomingPreview();
        }

        private void Options_Changed(object sender, RoutedEventArgs e)
        {
            UpdateUpcomingPreview();
        }

        private void UpdateSelectionState()
        {
            var selected = SelectedViewModel;
            if (selected == null)
            {
                EditButton.IsEnabled = false;
                DeleteButton.IsEnabled = false;
                ResetButton.IsEnabled = false;
                _upcomingPreview.Clear();
                return;
            }

            EditButton.IsEnabled = true;
            DeleteButton.IsEnabled = !selected.IsBuiltIn;
            ResetButton.IsEnabled = true;

            UpdateUpcomingPreview();
        }

        private void UpdateUpcomingPreview()
        {
            if (_upcomingPreview == null) return;
            _upcomingPreview.Clear();

            var selected = SelectedViewModel;
            if (selected == null) return;

            bool continueLast = ContinueFromLastCheck?.IsChecked == true;
            var samples = SequenceNumberEngine.GenerateSequence(
                selected.Definition,
                count: 6,
                continueFromCurrent: continueLast,
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
                _sequenceItems.Add(vm);
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
                selected.Definition.Prefix = updated.Prefix;
                selected.Definition.Suffix = updated.Suffix;
                selected.Definition.StartNumber = updated.StartNumber;
                selected.Definition.EndNumber = updated.EndNumber;
                selected.Definition.Increment = updated.Increment;
                selected.Definition.Digits = updated.Digits;
                selected.Definition.RepeatCount = updated.RepeatCount;

                selected.RefreshProperties();
                UpdateUpcomingPreview();
                SaveCurrentSequences();
                StatusLabel.Text = $"Updated sequence '{updated.Name}'.";
            }
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
                _sequenceItems.Remove(selected);
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
                    FillOrder = FillOrderCombo.SelectedIndex == 1 ? SequenceFillOrder.Horizontal : SequenceFillOrder.Vertical,
                    OnlyFillBlankCells = OnlyBlankCheck.IsChecked == true,
                    SkipHiddenCells = SkipHiddenCheck.IsChecked == true,
                    ContinueFromLast = ContinueFromLastCheck.IsChecked == true
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
                // Refresh target address label in case active selection moved
                string sheet = _context.ActiveWorksheetName ?? "ActiveSheet";
                string addr = _context.SelectionAddress ?? "Selection";
                SelectionAddressLabel.Text = $"{sheet}!{addr}";
            }
        }

        private void SaveCurrentSequences()
        {
            var defs = _sequenceItems.Select(vm => vm.Definition).ToList();
            SequenceNumberEngine.SaveSequences(defs);
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
