using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Navigation;

namespace ExcelProductivitySuite.UI.ViewModels
{
    /// <summary>
    /// ViewModel driving the Workbook and Sheet Navigator pane/dialog.
    /// Provides real-time synchronization, instant search/filtering, reordering, and workbook management.
    /// </summary>
    public class WorkbookSheetNavigatorViewModel : INotifyPropertyChanged
    {
        private readonly IExcelContext _context;
        private readonly List<WorkbookSummaryInfo> _rawWorkbooks = new();
        private readonly List<WorksheetSummaryInfo> _rawWorksheets = new();

        private WorkbookSummaryInfo? _selectedWorkbook;
        private WorksheetSummaryInfo? _selectedWorksheet;
        private string _searchText = string.Empty;
        private SheetFilterMode _filterMode = SheetFilterMode.All;
        private WorkbookSortOrder _workbookSort = WorkbookSortOrder.OpenOrder;
        private string _statusMessage = string.Empty;

        public event PropertyChangedEventHandler? PropertyChanged;

        public ObservableCollection<WorkbookSummaryInfo> Workbooks { get; } = new();
        public ObservableCollection<WorksheetSummaryInfo> Worksheets { get; } = new();

        public WorkbookSummaryInfo? SelectedWorkbook
        {
            get => _selectedWorkbook;
            set
            {
                if (_selectedWorkbook != value)
                {
                    _selectedWorkbook = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(HasSelectedWorkbook));
                    LoadWorksheetsForSelectedWorkbook();
                }
            }
        }

        public WorksheetSummaryInfo? SelectedWorksheet
        {
            get => _selectedWorksheet;
            set
            {
                if (_selectedWorksheet != value)
                {
                    _selectedWorksheet = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(HasSelectedWorksheet));
                }
            }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (_searchText != value)
                {
                    _searchText = value;
                    OnPropertyChanged();
                    ApplyWorksheetFilters();
                }
            }
        }

        public SheetFilterMode FilterMode
        {
            get => _filterMode;
            set
            {
                if (_filterMode != value)
                {
                    _filterMode = value;
                    OnPropertyChanged();
                    ApplyWorksheetFilters();
                }
            }
        }

        public WorkbookSortOrder WorkbookSort
        {
            get => _workbookSort;
            set
            {
                if (_workbookSort != value)
                {
                    _workbookSort = value;
                    OnPropertyChanged();
                    ApplyWorkbookSorting();
                }
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool HasSelectedWorkbook => SelectedWorkbook != null;
        public bool HasSelectedWorksheet => SelectedWorksheet != null;

        public string WorkbookCountSummary =>
            $"{Workbooks.Count} Open Workbook{(Workbooks.Count == 1 ? "" : "s")}";

        public string WorksheetCountSummary
        {
            get
            {
                int total = _rawWorksheets.Count;
                int visible = _rawWorksheets.Count(s => s.IsVisible);
                int hidden = total - visible;
                return $"{total} sheet{(total == 1 ? "" : "s")} ({visible} visible, {hidden} hidden)";
            }
        }

        public WorkbookSheetNavigatorViewModel(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            Refresh();
        }

        /// <summary>
        /// Reloads all open workbooks and active worksheets from Excel context.
        /// </summary>
        public void Refresh()
        {
            string? prevWbName = SelectedWorkbook?.Name;
            string? prevWsName = SelectedWorksheet?.Name;

            _rawWorkbooks.Clear();
            var wbList = _context.GetOpenWorkbookDetails();
            if (wbList != null)
            {
                _rawWorkbooks.AddRange(wbList);
            }

            ApplyWorkbookSorting();

            // Restore selection or select active workbook
            var matchWb = Workbooks.FirstOrDefault(w => string.Equals(w.Name, prevWbName, StringComparison.OrdinalIgnoreCase))
                       ?? Workbooks.FirstOrDefault(w => w.IsActive)
                       ?? Workbooks.FirstOrDefault();

            _selectedWorkbook = matchWb;
            OnPropertyChanged(nameof(SelectedWorkbook));
            OnPropertyChanged(nameof(HasSelectedWorkbook));
            LoadWorksheetsForSelectedWorkbook();

            if (SelectedWorkbook != null && !string.IsNullOrWhiteSpace(prevWsName))
            {
                var matchWs = Worksheets.FirstOrDefault(s => string.Equals(s.Name, prevWsName, StringComparison.OrdinalIgnoreCase));
                if (matchWs != null)
                {
                    SelectedWorksheet = matchWs;
                }
            }

            StatusMessage = $"Updated at {DateTime.Now:HH:mm:ss}";
            OnPropertyChanged(nameof(WorkbookCountSummary));
        }

        private void ApplyWorkbookSorting()
        {
            IEnumerable<WorkbookSummaryInfo> sorted = _workbookSort switch
            {
                WorkbookSortOrder.NameAscending => _rawWorkbooks.OrderBy(w => w.Name, StringComparer.OrdinalIgnoreCase),
                WorkbookSortOrder.NameDescending => _rawWorkbooks.OrderByDescending(w => w.Name, StringComparer.OrdinalIgnoreCase),
                _ => _rawWorkbooks
            };

            Workbooks.Clear();
            foreach (var wb in sorted)
            {
                Workbooks.Add(wb);
            }
            OnPropertyChanged(nameof(WorkbookCountSummary));
        }

        private void LoadWorksheetsForSelectedWorkbook()
        {
            _rawWorksheets.Clear();
            if (SelectedWorkbook != null)
            {
                var sheets = _context.GetWorksheetDetails(SelectedWorkbook.Name);
                if (sheets != null)
                {
                    _rawWorksheets.AddRange(sheets);
                }
            }

            ApplyWorksheetFilters();
            OnPropertyChanged(nameof(WorksheetCountSummary));
        }

        private void ApplyWorksheetFilters()
        {
            Worksheets.Clear();
            string query = SearchText?.Trim() ?? string.Empty;

            foreach (var ws in _rawWorksheets)
            {
                if (FilterMode == SheetFilterMode.VisibleOnly && !ws.IsVisible) continue;
                if (FilterMode == SheetFilterMode.HiddenOnly && ws.IsVisible) continue;

                if (!string.IsNullOrWhiteSpace(query))
                {
                    if (ws.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }
                }

                Worksheets.Add(ws);
            }

            // Restore selection if matching
            if (SelectedWorksheet != null)
            {
                SelectedWorksheet = Worksheets.FirstOrDefault(s => string.Equals(s.Name, SelectedWorksheet.Name, StringComparison.OrdinalIgnoreCase))
                                 ?? Worksheets.FirstOrDefault(s => s.IsActive)
                                 ?? Worksheets.FirstOrDefault();
            }
            else
            {
                SelectedWorksheet = Worksheets.FirstOrDefault(s => s.IsActive) ?? Worksheets.FirstOrDefault();
            }

            OnPropertyChanged(nameof(WorksheetCountSummary));
        }

        public void ActivateWorkbook(WorkbookSummaryInfo? wb = null)
        {
            var target = wb ?? SelectedWorkbook;
            if (target == null) return;

            bool ok = _context.ActivateWorkbook(target.Name);
            if (ok)
            {
                StatusMessage = $"Activated workbook: {target.Name}";
                Refresh();
            }
            else
            {
                StatusMessage = $"Could not activate workbook: {target.Name}";
            }
        }

        public void CloseWorkbook(WorkbookSummaryInfo? wb = null)
        {
            var target = wb ?? SelectedWorkbook;
            if (target == null) return;

            if (!target.IsSaved)
            {
                bool proceed = _context.Confirm(
                    $"Workbook '{target.Name}' has unsaved changes. Do you want to close it without saving?",
                    "Close Workbook");
                if (!proceed) return;
            }

            bool ok = _context.CloseWorkbook(target.Name, saveChanges: false);
            if (ok)
            {
                StatusMessage = $"Closed workbook: {target.Name}";
                Refresh();
            }
            else
            {
                StatusMessage = $"Could not close workbook: {target.Name}";
            }
        }

        public void SaveWorkbook(WorkbookSummaryInfo? wb = null)
        {
            var target = wb ?? SelectedWorkbook;
            if (target == null) return;

            bool ok = _context.SaveWorkbook(target.Name);
            if (ok)
            {
                StatusMessage = $"Saved workbook: {target.Name}";
                Refresh();
            }
            else
            {
                StatusMessage = $"Could not save workbook: {target.Name}";
            }
        }

        public void SaveAllWorkbooks()
        {
            _context.SaveAllWorkbooks();
            StatusMessage = "All open workbooks saved.";
            Refresh();
        }

        public void CreateNewWorkbook()
        {
            bool ok = _context.CreateNewWorkbook();
            if (ok)
            {
                StatusMessage = "Created new workbook.";
                Refresh();
            }
        }

        public void OpenWorkbookFolder(WorkbookSummaryInfo? wb = null)
        {
            var target = wb ?? SelectedWorkbook;
            if (target == null) return;

            if (!target.HasDiskPath || !File.Exists(target.Path))
            {
                _context.ShowMessage("This workbook has not yet been saved to disk.", "Open Containing Folder");
                return;
            }

            try
            {
                Process.Start("explorer.exe", $"/select,\"{target.Path}\"");
                StatusMessage = $"Opened folder for: {target.Name}";
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Failed to open folder: {ex.Message}", "Open Folder Error", isError: true);
            }
        }

        public void CopyWorkbookPath(WorkbookSummaryInfo? wb = null)
        {
            var target = wb ?? SelectedWorkbook;
            if (target == null) return;

            if (!target.HasDiskPath)
            {
                _context.ShowMessage("This workbook has not been saved to disk yet.", "Copy Path");
                return;
            }

            try
            {
                Clipboard.SetText(target.Path);
                StatusMessage = $"Copied path to clipboard: {target.Path}";
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Could not copy to clipboard: {ex.Message}", "Copy Path Error", isError: true);
            }
        }

        public void ActivateWorksheet(WorksheetSummaryInfo? ws = null)
        {
            var target = ws ?? SelectedWorksheet;
            if (target == null) return;

            bool ok = _context.ActivateWorksheet(target.Name, target.WorkbookName);
            if (ok)
            {
                StatusMessage = $"Switched to sheet: '{target.Name}'";
                Refresh();
            }
            else
            {
                StatusMessage = $"Could not switch to sheet: '{target.Name}'";
            }
        }

        public void ToggleLastActiveSheet()
        {
            bool ok = _context.ToggleLastActiveWorksheet(SelectedWorkbook?.Name);
            if (ok)
            {
                StatusMessage = "Toggled to previous sheet.";
                Refresh();
            }
            else
            {
                StatusMessage = "No previous sheet history found.";
            }
        }

        public void MoveWorksheet(SheetMoveDirection direction)
        {
            if (SelectedWorksheet == null || SelectedWorkbook == null) return;

            string name = SelectedWorksheet.Name;
            bool ok = _context.MoveWorksheet(name, direction, SelectedWorkbook.Name);
            if (ok)
            {
                StatusMessage = $"Moved '{name}' {direction.ToString().ToLowerInvariant()}.";
                Refresh();
            }
            else
            {
                StatusMessage = $"Could not move '{name}'.";
            }
        }

        public void ToggleSheetVisibility(WorksheetSummaryInfo? ws = null)
        {
            var target = ws ?? SelectedWorksheet;
            if (target == null || SelectedWorkbook == null) return;

            if (target.IsVisible && _rawWorksheets.Count(s => s.IsVisible) <= 1)
            {
                _context.ShowMessage("A workbook must contain at least one visible worksheet.", "Cannot Hide Sheet", isError: true);
                return;
            }

            bool makeVisible = target.IsHidden;
            _context.SetWorksheetVisibility(target.Name, makeVisible, veryHidden: false);
            StatusMessage = makeVisible ? $"Unhid sheet '{target.Name}'" : $"Hid sheet '{target.Name}'";
            Refresh();
        }

        public void AddNewWorksheet(string? customName = null)
        {
            if (SelectedWorkbook == null) return;

            string? validName = null;
            if (!string.IsNullOrWhiteSpace(customName))
            {
                if (!SheetNameValidator.Validate(customName, _rawWorksheets.Select(s => s.Name), out string err))
                {
                    _context.ShowMessage(err, "Invalid Sheet Name", isError: true);
                    return;
                }
                validName = customName!.Trim();
            }

            bool ok = _context.AddWorksheet(validName, SelectedWorkbook.Name);
            if (ok)
            {
                StatusMessage = validName != null ? $"Added worksheet '{validName}'" : "Added new worksheet";
                Refresh();
            }
            else
            {
                StatusMessage = "Could not add new worksheet.";
            }
        }

        public void DeleteWorksheet(WorksheetSummaryInfo? ws = null)
        {
            var target = ws ?? SelectedWorksheet;
            if (target == null || SelectedWorkbook == null) return;

            if (_rawWorksheets.Count <= 1)
            {
                _context.ShowMessage("A workbook must contain at least one worksheet.", "Cannot Delete Sheet", isError: true);
                return;
            }

            bool proceed = _context.Confirm(
                $"Are you sure you want to delete worksheet '{target.Name}'?\nThis action cannot be undone.",
                "Delete Worksheet");
            if (!proceed) return;

            bool ok = _context.DeleteWorksheet(target.Name, SelectedWorkbook.Name);
            if (ok)
            {
                StatusMessage = $"Deleted worksheet '{target.Name}'";
                Refresh();
            }
            else
            {
                StatusMessage = $"Could not delete worksheet '{target.Name}'";
            }
        }

        public void RenameWorksheet(string oldName, string newName)
        {
            if (SelectedWorkbook == null) return;

            if (!SheetNameValidator.Validate(newName, _rawWorksheets.Where(s => !string.Equals(s.Name, oldName, StringComparison.OrdinalIgnoreCase)).Select(s => s.Name), out string err))
            {
                _context.ShowMessage(err, "Invalid Sheet Name", isError: true);
                return;
            }

            bool ok = _context.RenameWorksheet(oldName, newName.Trim(), SelectedWorkbook.Name);
            if (ok)
            {
                StatusMessage = $"Renamed '{oldName}' to '{newName.Trim()}'";
                Refresh();
            }
            else
            {
                StatusMessage = $"Could not rename sheet '{oldName}'";
            }
        }

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
