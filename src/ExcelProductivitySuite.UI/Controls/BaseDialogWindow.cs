using System;
using System.Windows;
using System.Windows.Interop;
using ExcelProductivitySuite.Core.Abstractions;

namespace ExcelProductivitySuite.UI.Controls
{
    public class BaseDialogWindow : Window
    {
        public BaseDialogWindow()
        {
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ShowInTaskbar = false;
            ResizeMode = ResizeMode.NoResize;
            FontFamily = new System.Windows.Media.FontFamily("Segoe UI, Arial");
            FontSize = 13;
        }

        /// <summary>
        /// Displays this dialog modally parented to the host Excel application window.
        /// </summary>
        public bool? ShowModal(IExcelContext context)
        {
            if (context != null && context.MainWindowHandle != IntPtr.Zero)
            {
                var helper = new WindowInteropHelper(this);
                helper.Owner = context.MainWindowHandle;
            }

            return ShowDialog();
        }
    }
}
