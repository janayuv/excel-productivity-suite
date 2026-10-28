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
    public class ColorStatDisplayItem
    {
        public int ColorRgb { get; set; }
        public Brush SwatchBrush { get; set; } = Brushes.Transparent;
        public string ColorHex { get; set; } = "";
        public int CellCount { get; set; }
        public int NumericCount { get; set; }
        public double Sum { get; set; }
        public double Average { get; set; }
        public double? Min { get; set; }
        public double? Max { get; set; }

        public string SumFormatted => Sum.ToString("N2");
        public string AverageFormatted => Average.ToString("N2");
        public string MinFormatted => Min.HasValue ? Min.Value.ToString("N2") : "-";
        public string MaxFormatted => Max.HasValue ? Max.Value.ToString("N2") : "-";
    }

    public partial class CountByColorDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private IReadOnlyList<CellColorInfo> _cells = new List<CellColorInfo>();
        public ObservableCollection<ColorStatDisplayItem> StatItems { get; } = new();

        public CountByColorDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            StatsDataGrid.ItemsSource = StatItems;
            LoadCellColors();
        }

        private void LoadCellColors()
        {
            _cells = _context.GetSelectedCellColors();
            CalculateStats();
        }

        private void CalculateStats()
        {
            StatItems.Clear();
            bool isFill = FillColorRadio?.IsChecked == true;

            var stats = ColorAnalyzer.CalculateColorStatistics(_cells, isFill);

            foreach (var s in stats)
            {
                byte r = (byte)(s.ColorRgb & 0xFF);
                byte g = (byte)((s.ColorRgb >> 8) & 0xFF);
                byte b = (byte)((s.ColorRgb >> 16) & 0xFF);

                var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
                brush.Freeze();

                StatItems.Add(new ColorStatDisplayItem
                {
                    ColorRgb = s.ColorRgb,
                    SwatchBrush = brush,
                    ColorHex = $"#{r:X2}{g:X2}{b:X2}",
                    CellCount = s.CellCount,
                    NumericCount = s.NumericCount,
                    Sum = s.Sum,
                    Average = s.Average,
                    Min = s.Min,
                    Max = s.Max
                });
            }

            if (StatusLabel != null)
            {
                StatusLabel.Text = $"{StatItems.Count} color group(s) analyzed across {_cells.Count} cells.";
            }
        }

        private void ColorTarget_Changed(object sender, RoutedEventArgs e)
        {
            CalculateStats();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ExportSheet_Click(object sender, RoutedEventArgs e)
        {
            if (StatItems.Count == 0)
            {
                _context.ShowMessage("No statistics available to export.", "No Data", isError: true);
                return;
            }

            object?[,] matrix = new object?[StatItems.Count + 1, 7];
            matrix[0, 0] = "Color Hex";
            matrix[0, 1] = "Cell Count";
            matrix[0, 2] = "Numeric Count";
            matrix[0, 3] = "Sum";
            matrix[0, 4] = "Average";
            matrix[0, 5] = "Min";
            matrix[0, 6] = "Max";

            for (int i = 0; i < StatItems.Count; i++)
            {
                var item = StatItems[i];
                matrix[i + 1, 0] = item.ColorHex;
                matrix[i + 1, 1] = item.CellCount;
                matrix[i + 1, 2] = item.NumericCount;
                matrix[i + 1, 3] = item.Sum;
                matrix[i + 1, 4] = item.Average;
                matrix[i + 1, 5] = item.Min ?? 0;
                matrix[i + 1, 6] = item.Max ?? 0;
            }

            string sheetName = "Color_Summary";
            bool success = _context.CreateWorksheetWithData(sheetName, matrix, activate: true);

            if (success)
            {
                _context.ShowMessage($"Exported color statistics to worksheet '{sheetName}'.", "Export Complete");
                Close();
            }
        }
    }
}
