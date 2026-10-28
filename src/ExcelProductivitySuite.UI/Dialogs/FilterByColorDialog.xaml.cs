using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public class ColorSwatchItem
    {
        public int ColorRgb { get; set; }
        public Brush SwatchBrush { get; set; } = Brushes.Transparent;
        public string ColorHex { get; set; } = "";
        public int CellCount { get; set; }
        public string CountText => $"{CellCount} cell(s)";
    }

    public partial class FilterByColorDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private IReadOnlyList<CellColorInfo> _cells = new List<CellColorInfo>();
        public ObservableCollection<ColorSwatchItem> Swatches { get; } = new();

        public FilterByColorDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            ColorsListBox.ItemsSource = Swatches;
            LoadCellColors();
        }

        private void LoadCellColors()
        {
            _cells = _context.GetSelectedCellColors();
            RefreshSwatches();
        }

        private void RefreshSwatches()
        {
            Swatches.Clear();
            bool isFill = FillColorRadio?.IsChecked == true;

            var grouped = _cells
                .GroupBy(c => isFill ? c.FillColorRgb : c.FontColorRgb)
                .OrderByDescending(g => g.Count());

            foreach (var g in grouped)
            {
                int color = g.Key;
                byte r = (byte)(color & 0xFF);
                byte gr = (byte)((color >> 8) & 0xFF);
                byte b = (byte)((color >> 16) & 0xFF);

                var brush = new SolidColorBrush(Color.FromRgb(r, gr, b));
                brush.Freeze();

                Swatches.Add(new ColorSwatchItem
                {
                    ColorRgb = color,
                    SwatchBrush = brush,
                    ColorHex = $"RGB({r}, {gr}, {b}) [#{r:X2}{gr:X2}{b:X2}]",
                    CellCount = g.Count()
                });
            }

            if (StatusLabel != null)
            {
                StatusLabel.Text = $"{Swatches.Count} unique color(s) found across {_cells.Count} cells.";
            }

            if (SelectCellsButton != null)
            {
                SelectCellsButton.IsEnabled = false;
            }
        }

        private void ColorTarget_Changed(object sender, RoutedEventArgs e)
        {
            RefreshSwatches();
        }

        private void ColorsList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            SelectCellsButton.IsEnabled = ColorsListBox.SelectedItem is ColorSwatchItem;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void SelectMatching_Click(object sender, RoutedEventArgs e)
        {
            if (ColorsListBox.SelectedItem is ColorSwatchItem item)
            {
                bool isFill = FillColorRadio.IsChecked == true;
                var matches = ColorAnalyzer.FilterCellsByColor(_cells, item.ColorRgb, isFill);
                _context.SelectCells(matches);
                _context.ShowMessage($"Selected {matches.Count} cell(s) matching color {item.ColorHex}.", "Cells Selected");
                DialogResult = true;
                Close();
            }
        }
    }
}
