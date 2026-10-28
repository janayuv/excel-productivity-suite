using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Recovery;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class RestoreSnapshotDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private readonly string _snapshotDir;
        private readonly string? _activeWbName;

        public ObservableCollection<SnapshotFileInfo> Snapshots { get; } = new ObservableCollection<SnapshotFileInfo>();

        public RestoreSnapshotDialog(IExcelContext context)
        {
            InitializeComponent();
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Owner = Application.Current?.MainWindow;

            _snapshotDir = WorkbookSnapshotManager.GetDefaultSnapshotDirectory();
            _activeWbName = _context.ActiveWorkbookName;

            if (string.IsNullOrWhiteSpace(_activeWbName))
            {
                FilterActiveWbCheck.IsChecked = false;
                FilterActiveWbCheck.IsEnabled = false;
            }
            else
            {
                FilterActiveWbCheck.Content = $"Show only snapshots for '{_activeWbName}'";
            }

            SnapshotsGrid.ItemsSource = Snapshots;
            LoadSnapshots();
        }

        private void LoadSnapshots()
        {
            Snapshots.Clear();
            string? filter = FilterActiveWbCheck.IsChecked == true ? _activeWbName : null;
            var list = WorkbookSnapshotManager.GetSnapshots(_snapshotDir, filter);
            foreach (var item in list)
            {
                Snapshots.Add(item);
            }

            RestoreBtn.IsEnabled = false;
            DeleteBtn.IsEnabled = false;
            SelectedPathText.Text = "Selected: (None)";
        }

        private void Filter_Click(object sender, RoutedEventArgs e)
        {
            LoadSnapshots();
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            LoadSnapshots();
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

        private void SnapshotsGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (SnapshotsGrid.SelectedItem is SnapshotFileInfo selected)
            {
                RestoreBtn.IsEnabled = true;
                DeleteBtn.IsEnabled = true;
                SelectedPathText.Text = $"Selected: {selected.FilePath}";
            }
            else
            {
                RestoreBtn.IsEnabled = false;
                DeleteBtn.IsEnabled = false;
                SelectedPathText.Text = "Selected: (None)";
            }
        }

        private void Restore_Click(object sender, RoutedEventArgs e)
        {
            if (SnapshotsGrid.SelectedItem is not SnapshotFileInfo selected) return;

            bool confirm = _context.Confirm(
                $"Are you sure you want to restore the snapshot '{selected.FileName}'?\n\nThis will close the active workbook without saving and load the snapshot.",
                "Confirm Snapshot Restore");

            if (!confirm) return;

            bool success = _context.RestoreWorkbookFromSnapshot(selected.FilePath);
            if (success)
            {
                _context.ShowMessage($"Snapshot '{selected.FileName}' restored successfully.", "Restore Complete");
                DialogResult = true;
                Close();
            }
            else
            {
                _context.ShowMessage("Failed to restore snapshot. Please verify file access permissions.", "Restore Error", isError: true);
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            if (SnapshotsGrid.SelectedItem is not SnapshotFileInfo selected) return;

            bool confirm = _context.Confirm(
                $"Are you sure you want to delete snapshot '{selected.FileName}'?\nThis action cannot be undone.",
                "Confirm Delete");

            if (!confirm) return;

            bool deleted = WorkbookSnapshotManager.DeleteSnapshot(selected.FilePath);
            if (deleted)
            {
                LoadSnapshots();
            }
            else
            {
                _context.ShowMessage("Failed to delete snapshot file.", "Delete Error", isError: true);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
