using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Navigation;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class DatePickerDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private DateTime _selectedDate = DateTime.Today;

        public DatePickerDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            TargetCellText.Text = $"{_context.SelectionAddress ?? "$A$1"} ({_context.ActiveWorksheetName ?? "Active Sheet"})";

            CalendarControl.SelectedDate = DateTime.Today;

            FormatCombo.ItemsSource = DateFormatPresets.GetStandardPresets(DateTime.Today);
            FormatCombo.SelectedIndex = 0;

            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewText == null || CustomFormatBox == null) return;
            string pattern = CustomFormatBox.Text.Trim();
            if (string.IsNullOrEmpty(pattern)) pattern = "yyyy-MM-dd";

            string formatted = DateFormatPresets.Format(_selectedDate, pattern);
            PreviewText.Text = formatted;
        }

        private void Calendar_SelectedDatesChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (CalendarControl.SelectedDate.HasValue)
            {
                _selectedDate = CalendarControl.SelectedDate.Value;
                UpdatePreview();
            }
        }

        private void Today_Click(object sender, RoutedEventArgs e)
        {
            CalendarControl.SelectedDate = DateTime.Today;
        }

        private void FormatCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (FormatCombo.SelectedItem is DateFormatOption opt)
            {
                CustomFormatBox.Text = opt.FormatString;
                UpdatePreview();
            }
        }

        private void CustomFormat_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void InsertDate_Click(object sender, RoutedEventArgs e)
        {
            string valueToInsert = PreviewText.Text;

            // Write formatted date string into the selection
            var matrix = new object?[1, 1] { { valueToInsert } };
            _context.SetSelectedValues(matrix);

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
