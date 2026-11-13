using System.Windows;
using ExcelProductivitySuite.Core;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class AboutDialog : BaseDialogWindow
    {
        private readonly IExcelContext? _context;

        public AboutDialog(IExcelContext? context = null)
        {
            _context = context;
            InitializeComponent();
            VersionLabel.Text = $"Version {ProductInfo.CurrentVersion} • {ProductInfo.Edition}";
        }

        private void CheckUpdatesButton_Click(object sender, RoutedEventArgs e)
        {
            var updateDlg = new CheckForUpdatesDialog(_context);
            if (_context != null)
            {
                updateDlg.ShowModal(_context);
            }
            else
            {
                updateDlg.ShowDialog();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
