using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
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
        private TextBox? _lastFocusedTextBox;
        private bool _isInitialized;

        public SequenceNumberDefinition ResultSequence { get; private set; }

        public CreateSequenceNumberDialog() : this(null, false)
        {
        }

        public CreateSequenceNumberDialog(SequenceNumberDefinition? definition, bool isEditMode = false)
        {
            InitializeComponent();
            _isEditMode = isEditMode;
            PreviewList.ItemsSource = _previewItems;

            PopulateSequenceTypes();

            if (_isEditMode && definition != null)
            {
                Title = "Edit Sequence Number - Excel Productivity Suite";
                DialogTitleText.Text = "Edit Sequence Number";
                _originalId = definition.Id;
                _isBuiltIn = definition.IsBuiltIn;
                _currentNumber = definition.CurrentNumber;

                NameText.Text = definition.Name;
                SetSelectedSequenceType(definition.Type);

                if (definition.CustomListItems != null && definition.CustomListItems.Count > 0)
                {
                    CustomItemsText.Text = string.Join(", ", definition.CustomListItems);
                }

                PrefixText.Text = definition.Prefix;
                StartNumberText.Text = definition.StartNumber.ToString();
                IncrementText.Text = definition.Increment.ToString();
                DigitsText.Text = definition.Digits.ToString();
                SuffixText.Text = definition.Suffix;
                RepeatCountText.Text = definition.RepeatCount.ToString();
                UsePlaceholdersCheck.IsChecked = definition.UseDynamicPlaceholders;
                OutputNumericCheck.IsChecked = definition.OutputAsNumeric;

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
            else
            {
                SequenceTypeCombo.SelectedIndex = 0;
            }

            ResultSequence = definition?.Clone() ?? new SequenceNumberDefinition();
            _isInitialized = true;
            UpdateLivePreview();
        }

        private void PopulateSequenceTypes()
        {
            if (SequenceTypeCombo == null) return;
            SequenceTypeCombo.Items.Clear();
            SequenceTypeCombo.Items.Add(new ComboBoxItem { Content = "Numeric (1, 2, 3...)", Tag = SequenceType.Numeric });
            SequenceTypeCombo.Items.Add(new ComboBoxItem { Content = "Uppercase Letters (A, B, C...)", Tag = SequenceType.AlphaUpper });
            SequenceTypeCombo.Items.Add(new ComboBoxItem { Content = "Lowercase Letters (a, b, c...)", Tag = SequenceType.AlphaLower });
            SequenceTypeCombo.Items.Add(new ComboBoxItem { Content = "Uppercase Roman (I, II, III...)", Tag = SequenceType.RomanUpper });
            SequenceTypeCombo.Items.Add(new ComboBoxItem { Content = "Lowercase Roman (i, ii, iii...)", Tag = SequenceType.RomanLower });
            SequenceTypeCombo.Items.Add(new ComboBoxItem { Content = "Hexadecimal (0001, 0002...)", Tag = SequenceType.Hexadecimal });
            SequenceTypeCombo.Items.Add(new ComboBoxItem { Content = "Custom List (Cycles through items)", Tag = SequenceType.CustomList });
        }

        private void SetSelectedSequenceType(SequenceType type)
        {
            if (SequenceTypeCombo == null) return;
            foreach (ComboBoxItem item in SequenceTypeCombo.Items)
            {
                if (item.Tag is SequenceType st && st == type)
                {
                    SequenceTypeCombo.SelectedItem = item;
                    break;
                }
            }
        }

        private SequenceType GetSelectedSequenceType()
        {
            if (SequenceTypeCombo?.SelectedItem is ComboBoxItem item && item.Tag is SequenceType st)
            {
                return st;
            }
            return SequenceType.Numeric;
        }

        private void SequenceTypeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;
            var type = GetSelectedSequenceType();

            bool isCustom = type == SequenceType.CustomList;
            if (CustomItemsLabel != null) CustomItemsLabel.Visibility = isCustom ? Visibility.Visible : Visibility.Collapsed;
            if (CustomItemsText != null) CustomItemsText.Visibility = isCustom ? Visibility.Visible : Visibility.Collapsed;

            if (OutputNumericCheck != null)
            {
                OutputNumericCheck.IsEnabled = type == SequenceType.Numeric;
                if (!OutputNumericCheck.IsEnabled) OutputNumericCheck.IsChecked = false;
            }

            UpdateLivePreview();
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb)
            {
                _lastFocusedTextBox = tb;
            }
        }

        private void InsertPlaceholder_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string token)
            {
                var target = _lastFocusedTextBox ?? PrefixText;
                if (target != null)
                {
                    int caret = target.SelectionStart;
                    string current = target.Text ?? "";
                    target.Text = current.Insert(caret, token);
                    target.SelectionStart = caret + token.Length;
                    target.Focus();
                }
            }
        }

        private void EndNumberCheck_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            if (EndNumberText == null) return;
            EndNumberText.IsEnabled = EndNumberCheck.IsChecked == true;
            UpdateLivePreview();
        }

        private void Input_Changed(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdateLivePreview();
        }

        private void OutputOptions_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) return;
            UpdateLivePreview();
        }

        private void UpdateLivePreview()
        {
            if (!_isInitialized || _previewItems == null) return;

            var def = BuildDefinitionFromInputs(out string? error);
            _previewItems.Clear();

            if (def != null)
            {
                var samples = def.GeneratePreview(8, DateTime.Now, "Sheet1");
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

            var seqType = GetSelectedSequenceType();

            // Parse Start Number
            string startRaw = StartNumberText?.Text?.Trim() ?? "1";
            long startNum = 1;

            if (seqType == SequenceType.AlphaUpper || seqType == SequenceType.AlphaLower)
            {
                if (!long.TryParse(startRaw, out startNum))
                {
                    startNum = SequenceNumberDefinition.AlphaToNumber(startRaw);
                }
            }
            else if (seqType == SequenceType.RomanUpper || seqType == SequenceType.RomanLower)
            {
                if (!long.TryParse(startRaw, out startNum))
                {
                    startNum = SequenceNumberDefinition.RomanToNumber(startRaw);
                }
            }
            else if (seqType == SequenceType.Hexadecimal)
            {
                if (startRaw.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
                {
                    startRaw = startRaw.Substring(2);
                }

                if (!long.TryParse(startRaw, System.Globalization.NumberStyles.HexNumber, null, out startNum))
                {
                    if (!long.TryParse(startRaw, out startNum))
                    {
                        errorMessage = "Start value must be a valid integer or hex code.";
                        return null;
                    }
                }
            }
            else
            {
                if (!long.TryParse(startRaw, out startNum))
                {
                    errorMessage = "Start number must be a valid integer.";
                    return null;
                }
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
                string endRaw = EndNumberText?.Text?.Trim() ?? "";
                if (!long.TryParse(endRaw, out long en))
                {
                    if (seqType == SequenceType.AlphaUpper || seqType == SequenceType.AlphaLower)
                        en = SequenceNumberDefinition.AlphaToNumber(endRaw);
                    else if (seqType == SequenceType.RomanUpper || seqType == SequenceType.RomanLower)
                        en = SequenceNumberDefinition.RomanToNumber(endRaw);
                    else
                    {
                        errorMessage = "End limit must be a valid integer.";
                        return null;
                    }
                }

                if (increment > 0 && en <= startNum)
                {
                    errorMessage = "End limit must be greater than start value when increment is positive.";
                    return null;
                }
                else if (increment < 0 && en >= startNum)
                {
                    errorMessage = "End limit must be less than start value when increment is negative.";
                    return null;
                }

                endNum = en;
            }

            List<string> customItems = new();
            if (seqType == SequenceType.CustomList)
            {
                string rawCustom = CustomItemsText?.Text ?? "";
                var parts = rawCustom.Split(new[] { ',', '\r', '\n', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                     .Select(p => p.Trim())
                                     .Where(p => !string.IsNullOrEmpty(p))
                                     .ToList();
                if (parts.Count == 0)
                {
                    errorMessage = "Custom list items must contain at least one element.";
                    return null;
                }
                customItems = parts;
            }

            return new SequenceNumberDefinition
            {
                Id = !string.IsNullOrWhiteSpace(_originalId) ? _originalId : Guid.NewGuid().ToString("N"),
                Name = name,
                Type = seqType,
                CustomListItems = customItems,
                Prefix = PrefixText?.Text ?? "",
                Suffix = SuffixText?.Text ?? "",
                StartNumber = startNum,
                EndNumber = endNum,
                Increment = increment,
                Digits = digits,
                RepeatCount = repeatCount,
                CurrentNumber = _isEditMode ? _currentNumber : startNum,
                IsBuiltIn = _isBuiltIn,
                UseDynamicPlaceholders = UsePlaceholdersCheck?.IsChecked == true,
                OutputAsNumeric = OutputNumericCheck?.IsChecked == true
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
