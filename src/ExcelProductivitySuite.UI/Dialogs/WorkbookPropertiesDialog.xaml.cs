using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class WorkbookPropertiesDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;

        public WorkbookPropertiesDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();
            LoadProperties();
        }

        private void LoadProperties()
        {
            string? name = _context.ActiveWorkbookName;
            string? path = _context.ActiveWorkbookPath;

            WbNameText.Text = string.IsNullOrEmpty(name) ? "(No active workbook)" : name;
            WbPathText.Text = string.IsNullOrEmpty(path) ? "(Unsaved or in-memory workbook)" : path;

            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                var fi = new FileInfo(path);
                FileSizeText.Text = FormatFileSize(fi.Length);
                LastModifiedText.Text = fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
            else
            {
                FileSizeText.Text = "Not saved on disk";
                LastModifiedText.Text = "N/A";
            }

            ActiveSheetText.Text = _context.ActiveWorksheetName ?? "None";

            var visibilities = _context.GetWorksheetVisibilities();
            if (visibilities != null && visibilities.Count > 0)
            {
                int total = visibilities.Count;
                int visible = visibilities.Count(v => v.IsVisible && !v.IsVeryHidden);
                int hidden = total - visible;
                SheetCountsText.Text = $"{total} sheet(s) ({visible} visible, {hidden} hidden)";
            }
            else
            {
                var names = _context.GetWorksheetNames();
                int count = names?.Count ?? 0;
                SheetCountsText.Text = $"{count} sheet(s)";
            }

            WbProtectionText.Text = _context.IsWorkbookProtected ? "Protected (Structure locked)" : "Unprotected";
            SheetProtectionText.Text = _context.IsCurrentSheetProtected ? "Protected" : "Unprotected";
        }

        private static string FormatFileSize(long bytes)
        {
            if (bytes >= 1024 * 1024)
            {
                return $"{(bytes / (1024.0 * 1024.0)):F2} MB ({bytes:N0} bytes)";
            }
            if (bytes >= 1024)
            {
                return $"{(bytes / 1024.0):F2} KB ({bytes:N0} bytes)";
            }
            return $"{bytes} bytes";
        }

        private void CopyPathButton_Click(object sender, RoutedEventArgs e)
        {
            string? path = _context.ActiveWorkbookPath;
            if (string.IsNullOrEmpty(path))
            {
                _context.ShowMessage("The workbook has not been saved to disk yet.", "Workbook Path");
                return;
            }

            try
            {
                Clipboard.SetText(path);
                _context.ShowMessage($"Copied path to clipboard:\n\n{path}", "Workbook Path Copied");
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Could not copy path: {ex.Message}", "Clipboard Error", isError: true);
            }
        }

        private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            string? path = _context.ActiveWorkbookPath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                _context.ShowMessage("The workbook has not been saved to disk yet or does not exist on disk.", "Open Folder", isError: true);
                return;
            }

            try
            {
                Process.Start("explorer.exe", $"/select,\"{path}\"");
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Could not open folder: {ex.Message}", "Error", isError: true);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
