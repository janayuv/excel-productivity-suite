using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formulas;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class FormulaTemplatesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly Dictionary<string, TextBox> _paramInputs = new Dictionary<string, TextBox>();
        private FormulaTemplate? _selectedTemplate;

        public FormulaTemplatesDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            InitCategoryCombo();
            FilterTemplates();
        }

        private void InitCategoryCombo()
        {
            CategoryCombo.Items.Clear();
            CategoryCombo.Items.Add("All Categories");
            CategoryCombo.Items.Add("Lookup & Reference");
            CategoryCombo.Items.Add("Financial & Business");
            CategoryCombo.Items.Add("Math & Statistics");
            CategoryCombo.Items.Add("Text Manipulation");
            CategoryCombo.Items.Add("Date & Time");
            CategoryCombo.Items.Add("Logical & Conditional");
            CategoryCombo.SelectedIndex = 0;
        }

        private FormulaCategory GetSelectedCategory()
        {
            return CategoryCombo.SelectedIndex switch
            {
                1 => FormulaCategory.LookupAndReference,
                2 => FormulaCategory.Financial,
                3 => FormulaCategory.MathAndStatistics,
                4 => FormulaCategory.TextManipulation,
                5 => FormulaCategory.DateAndTime,
                6 => FormulaCategory.LogicalAndConditional,
                _ => FormulaCategory.All
            };
        }

        private void FilterTemplates()
        {
            var cat = GetSelectedCategory();
            string query = SearchBox.Text;
            var list = FormulaTemplateLibrary.Search(query, cat);

            TemplatesList.ItemsSource = list;
            if (list.Count > 0)
            {
                TemplatesList.SelectedIndex = 0;
            }
            else
            {
                ClearTemplateDetails();
            }
        }

        private void CategoryCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (IsLoaded) FilterTemplates();
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (IsLoaded) FilterTemplates();
        }

        private void TemplatesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TemplatesList.SelectedItem is FormulaTemplate template)
            {
                DisplayTemplate(template);
            }
        }

        private void DisplayTemplate(FormulaTemplate template)
        {
            _selectedTemplate = template;
            TemplateTitleText.Text = template.Name;
            TemplateDescText.Text = template.Description;
            TemplatePatternText.Text = template.TemplatePattern;

            ParametersPanel.Children.Clear();
            _paramInputs.Clear();

            string currentSelection = _context.SelectionAddress ?? "$A$1";

            foreach (var p in template.Parameters)
            {
                var rowGrid = new Grid { Margin = new Thickness(0, 4, 0, 8) };
                rowGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                rowGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                rowGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                var labelDock = new DockPanel();
                var label = new TextBlock
                {
                    Text = p.DisplayName + ":",
                    FontWeight = FontWeights.SemiBold,
                    FontSize = 11
                };
                labelDock.Children.Add(label);

                var useSelBtn = new Button
                {
                    Content = "Use Selection",
                    FontSize = 10,
                    Padding = new Thickness(4, 1, 4, 1),
                    Margin = new Thickness(6, 0, 0, 0),
                    Style = (Style)FindResource("ModernSecondaryButtonStyle")
                };
                DockPanel.SetDock(useSelBtn, Dock.Right);

                var textBox = new TextBox
                {
                    Text = p.DefaultValue,
                    FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                    Style = (Style)FindResource("ModernTextBoxStyle"),
                    Margin = new Thickness(0, 2, 0, 2)
                };

                useSelBtn.Click += (s, e) =>
                {
                    textBox.Text = currentSelection;
                };

                labelDock.Children.Add(useSelBtn);
                Grid.SetRow(labelDock, 0);
                Grid.SetRow(textBox, 1);

                var helpText = new TextBlock
                {
                    Text = p.Description,
                    FontSize = 10,
                    Foreground = (System.Windows.Media.Brush)FindResource("TextSecondaryBrush")
                };
                Grid.SetRow(helpText, 2);

                rowGrid.Children.Add(labelDock);
                rowGrid.Children.Add(textBox);
                rowGrid.Children.Add(helpText);

                textBox.TextChanged += (s, e) => UpdatePreview();
                _paramInputs[p.Key] = textBox;
                ParametersPanel.Children.Add(rowGrid);
            }

            UpdatePreview();
            InsertBtn.IsEnabled = true;
            StatusText.Text = $"Configuring {template.Name}. Click 'Insert into Active Cell' when ready.";
        }

        private void ClearTemplateDetails()
        {
            _selectedTemplate = null;
            TemplateTitleText.Text = "No templates match filter";
            TemplateDescText.Text = "Try altering your category or search query.";
            TemplatePatternText.Text = "";
            ParametersPanel.Children.Clear();
            _paramInputs.Clear();
            PreviewFormulaBox.Text = "";
            InsertBtn.IsEnabled = false;
        }

        private void UpdatePreview()
        {
            if (_selectedTemplate == null)
            {
                PreviewFormulaBox.Text = "";
                return;
            }

            var dict = new Dictionary<string, string>();
            foreach (var kvp in _paramInputs)
            {
                dict[kvp.Key] = kvp.Value.Text;
            }

            PreviewFormulaBox.Text = _selectedTemplate.RenderFormula(dict);
        }

        private void CopyPreviewBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(PreviewFormulaBox.Text))
            {
                try
                {
                    Clipboard.SetText(PreviewFormulaBox.Text);
                    StatusText.Text = "Formula copied to clipboard!";
                }
                catch { }
            }
        }

        private void InsertBtn_Click(object sender, RoutedEventArgs e)
        {
            string formula = PreviewFormulaBox.Text;
            if (string.IsNullOrWhiteSpace(formula)) return;

            try
            {
                _context.InsertFormulaAtActiveCell(formula);
                _context.ShowMessage($"Inserted template formula into active cell:\n{formula}", "Template Inserted");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Error inserting formula: {ex.Message}", "Insertion Error", isError: true);
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
