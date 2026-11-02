using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.AddIn.Navigation
{
    /// <summary>
    /// Windows Forms UserControl hosting the WPF WorkbookSheetNavigatorControl inside a VSTO CustomTaskPane.
    /// </summary>
    public class WorkbookSheetNavigatorTaskPaneControl : UserControl
    {
        private readonly ElementHost _elementHost;
        public WorkbookSheetNavigatorControl WpfNavigator { get; }

        public WorkbookSheetNavigatorTaskPaneControl(IExcelContext context)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));

            DoubleBuffered = true;
            BackColor = Color.White;
            Dock = DockStyle.Fill;

            WpfNavigator = new WorkbookSheetNavigatorControl(context, isTaskPane: true);
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = WpfNavigator
            };

            Controls.Add(_elementHost);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _elementHost?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
