using System;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class WorkbookSheetNavigatorWindow : BaseDialogWindow
    {
        public WorkbookSheetNavigatorControl NavigatorControl { get; }

        public WorkbookSheetNavigatorWindow(IExcelContext context)
        {
            InitializeComponent();
            if (context == null) throw new ArgumentNullException(nameof(context));

            NavigatorControl = new WorkbookSheetNavigatorControl(context, isTaskPane: false);
            HostGrid.Children.Add(NavigatorControl);
            Owner = Application.Current?.MainWindow;
        }
    }
}
