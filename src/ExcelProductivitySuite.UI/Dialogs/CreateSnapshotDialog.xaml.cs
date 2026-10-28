using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Recovery;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class CreateSnapshotDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly string _snapshotDir;
        private readonly string _wbName;

        public CreateSnapshotDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            _wbName = _context.ActiveWorkbookName ?? "Workbook.xlsx";
            WbNameText.Text = _wbName;
            WbPathText.Text = _context.ActiveWorkbookPath ?? "(Unsaved / Temporary Workbook)";

            _snapshotDir = WorkbookSnapshotManager.GetDefaultSnapshotDirectory();
            DirBox.Text = _snapshotDir;

            UpdatePreview();
        }

        private void UpdatePreview()
        {
            string tag = TagBox.Text.Trim();
            string path = WorkbookSnapshotManager.GenerateSnapshotFilePath(_snapshotDir, _wbName, string.IsNullOrEmpty(tag) ? null : tag);
            PreviewNameText.Text = Path.GetFileName(path);
        }

        private void TagBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Directory.Exists(_snapshotDir))
                {
                    Process.Start("explorer.exe", _snapshotDir);
                }
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Could not open folder: {ex.Message}", "Explorer Error", isError: true);
            }
        }

        private void TakeSnapshot_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string tag = TagBox.Text.Trim();
                string destinationPath = WorkbookSnapshotManager.GenerateSnapshotFilePath(_snapshotDir, _wbName, string.IsNullOrEmpty(tag) ? null : tag);

                bool success = _context.SaveActiveWorkbookCopy(destinationPath);
                if (success)
                {
                    _context.ShowMessage(
                        $"Workbook snapshot captured successfully!\n\nFile: {Path.GetFileName(destinationPath)}\nLocation: {_snapshotDir}\n\nYou can restore this state at any time using 'Restore Snapshot'.",
                        "Snapshot Created");
                    DialogResult = true;
                    Close();
                }
                else
                {
                    _context.ShowMessage("Failed to create snapshot. Please make sure the active workbook is saved and accessible.", "Snapshot Failed", isError: true);
                }
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Error creating snapshot: {ex.Message}", "Snapshot Error", isError: true);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
