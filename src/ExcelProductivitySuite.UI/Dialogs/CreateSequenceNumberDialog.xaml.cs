using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class CreateSequenceNumberDialog : BaseDialogWindow
    {
        private readonly ObservableCollection<string> _previewItems = new();
        private readonly bool _isEditMode;
        private string _originalId = string.Empty;
        private bool _isBuiltIn;
        private long _currentNumber = 1;

        public SequenceNumberDefinition ResultSequence { get; private set; }

        public CreateSequenceNumberDialog() : this(null, false)
        {
        }

        public CreateSequenceNumberDialog(SequenceNumberDefinition? definition, bool isEditMode = false)
        {
            InitializeComponent();
            _isEditMode = isEditMode;
            PreviewList.ItemsSource = _previewItems;

            if (_isEditMode && definition != null)
            {
                Title = "Edit Sequence Number - Excel Productivity Suite";
                DialogTitleText.Text = "Edit Sequence Number";
                _originalId = definition.Id;
                _isBuiltIn = definition.IsBuiltIn;
                _currentNumber = definition.CurrentNumber;

                NameText.Text = definition.Name;
                PrefixText.Text = definition.Prefix;
                StartNumberText.Text = definition.StartNumber.ToString();
                IncrementText.Text = definition.Increment.ToString();
                DigitsText.Text = definition.Digits.ToString();
                SuffixText.Text = definition.Suffix;
                RepeatCountText.Text = definition.RepeatCount.ToString();

                if (definition.EndNumber.HasValue)
                {
                    EndNumberCheck.IsChecked = true;
                    EndNumberText.Text = definition.EndNumber.Value.ToString();
                    EndNumberText.IsEnabled = true;
                }
                else
                {
                    EndNumberCheck.IsChecked = false;
                    EndNumberText.IsEnabled = false;
                }
            }

            ResultSequence = definition?.Clone() ?? new SequenceNumberDefinition();
            UpdateLivePreview();
        }

        private void EndNumberCheck_Changed(object sender, RoutedEventArgs e)
        {
            if (EndNumberText == null) return;
            EndNumberText.IsEnabled = EndNumberCheck.IsChecked == true;
            UpdateLivePreview();
        }

        private void Input_Changed(object sender, TextChangedEventArgs e)
        {
            UpdateLivePreview();
        }

        private void UpdateLivePreview()
        {
            if (_previewItems == null) return;

            var def = BuildDefinitionFromInputs(out string? error);
            _previewItems.Clear();

            if (def != null)
            {
                var samples = def.GeneratePreview(8);
                foreach (var s in samples)
                {
                    _previewItems.Add(s);
                }
                if (ValidationStatusLabel != null) ValidationStatusLabel.Text = "";
            }
            else
            {
                if (ValidationStatusLabel != null && !string.IsNullOrWhiteSpace(error))
                {
                    ValidationStatusLabel.Text = error;
                }
            }
        }

        private SequenceNumberDefinition? BuildDefinitionFromInputs(out string? errorMessage)
        {
            errorMessage = null;

            string name = NameText?.Text?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(name))
            {
                errorMessage = "Sequence name is required.";
                return null;
            }

            if (!long.TryParse(StartNumberText?.Text?.Trim(), out long startNum))
            {
                errorMessage = "Start number must be a valid integer.";
                return null;
            }

            if (!long.TryParse(IncrementText?.Text?.Trim(), out long increment) || increment == 0)
            {
                errorMessage = "Increment must be a non-zero integer.";
                return null;
            }

            if (!int.TryParse(DigitsText?.Text?.Trim(), out int digits) || digits < 1 || digits > 12)
            {
                errorMessage = "No. of digits must be between 1 and 12.";
                return null;
            }

            if (!int.TryParse(RepeatCountText?.Text?.Trim(), out int repeatCount) || repeatCount < 1)
            {
                errorMessage = "Repeat count must be 1 or greater.";
                return null;
            }

            long? endNum = null;
            if (EndNumberCheck?.IsChecked == true)
            {
                if (!long.TryParse(EndNumberText?.Text?.Trim(), out long en))
                {
                    errorMessage = "End number must be a valid integer.";
                    return null;
                }

                if (increment > 0 && en <= startNum)
                {
                    errorMessage = "End number must be greater than start number when increment is positive.";
                    return null;
                }
                else if (increment < 0 && en >= startNum)
                {
                    errorMessage = "End number must be less than start number when increment is negative.";
                    return null;
                }

                endNum = en;
            }

            return new SequenceNumberDefinition
            {
                Id = !string.IsNullOrWhiteSpace(_originalId) ? _originalId : Guid.NewGuid().ToString("N"),
                Name = name,
                Prefix = PrefixText?.Text ?? "",
                Suffix = SuffixText?.Text ?? "",
                StartNumber = startNum,
                EndNumber = endNum,
                Increment = increment,
                Digits = digits,
                RepeatCount = repeatCount,
                CurrentNumber = _isEditMode ? _currentNumber : startNum,
                IsBuiltIn = _isBuiltIn
            };
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var def = BuildDefinitionFromInputs(out string? error);
            if (def == null)
            {
                ValidationStatusLabel.Text = error ?? "Please correct the invalid input fields.";
                return;
            }

            ResultSequence = def;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
