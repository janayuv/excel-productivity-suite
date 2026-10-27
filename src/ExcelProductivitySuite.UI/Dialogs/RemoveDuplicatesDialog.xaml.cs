using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class RemoveDuplicatesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public RemoveDuplicatesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            var matrix = _context.GetSelectedValues();
            if (matrix == null)
            {
                _context.ShowMessage("No cells selected.", "Remove Duplicates", isError: true);
                return;
            }

            var options = new DuplicateDetectionOptions
            {
                KeepStrategy = KeepFirstRadio.IsChecked == true ? DuplicateKeepStrategy.KeepFirst : DuplicateKeepStrategy.KeepLast,
                MatchCase = MatchCaseCheckBox.IsChecked == true,
                CompareEntireRows = true
            };

            bool[] isDuplicate = DuplicateFinder.IdentifyDuplicateRows(matrix, options);

            int dupCount = 0;
            foreach (bool dup in isDuplicate)
            {
                if (dup) dupCount++;
            }

            if (dupCount == 0)
            {
                _context.ShowMessage("No duplicate rows were found in the selected range.", "Remove Duplicates");
                DialogResult = true;
                Close();
                return;
            }

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            var updated = (object?[,])matrix.Clone();

            for (int r = rStart; r <= rEnd; r++)
            {
                int rowIndex = r - rStart;
                if (isDuplicate[rowIndex])
                {
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        updated[r, c] = "";
                    }
                }
            }

            _context.SetSelectedValues(updated);
            _context.ShowMessage($"Successfully cleared {dupCount} duplicate rows.", "Remove Duplicates");

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
