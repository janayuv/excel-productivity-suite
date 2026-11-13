using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using ExcelProductivitySuite.Core;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Updates;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class CheckForUpdatesDialog : BaseDialogWindow
    {
        private readonly IExcelContext? _context;
        private readonly IUpdateService _updateService;
        private UpdateInfo? _currentUpdateInfo;

        public FrameworkElement CheckingStatePanel => CheckingPanel;
        public FrameworkElement UpdateAvailableStatePanel => UpdateAvailablePanel;
        public FrameworkElement DownloadingStatePanel => DownloadingPanel;
        public FrameworkElement UpToDateStatePanel => UpToDatePanel;
        public FrameworkElement ErrorStatePanel => ErrorPanel;

        public CheckForUpdatesDialog() : this(null, null)
        {
        }

        public CheckForUpdatesDialog(IExcelContext? context) : this(context, null)
        {
        }

        public CheckForUpdatesDialog(IExcelContext? context, IUpdateService? updateService)
        {
            _context = context;
            _updateService = updateService ?? new GitHubUpdateService();

            InitializeComponent();
            Loaded += async (s, e) => await CheckForUpdatesAsync();
        }

        public async Task CheckForUpdatesAsync()
        {
            ShowState(DialogState.Checking);

            try
            {
                var update = await _updateService.CheckForUpdatesAsync();
                _currentUpdateInfo = update;

                if (!string.IsNullOrEmpty(update.ErrorMessage))
                {
                    ErrorMessageText.Text = update.ErrorMessage;
                    ShowState(DialogState.Error);
                }
                else if (update.IsUpdateAvailable)
                {
                    VersionComparisonText.Text = $"Current: v{update.CurrentVersion}  ➔  Latest: v{update.LatestVersion}";
                    ReleaseNotesText.Text = string.IsNullOrWhiteSpace(update.ReleaseNotes) ? "No release notes provided." : update.ReleaseNotes;

                    string dateStr = update.PublishedAt.HasValue
                        ? update.PublishedAt.Value.ToString("MMM dd, yyyy")
                        : "Latest Release";
                    ReleaseDateText.Text = $"Released: {dateStr}";

                    ShowState(DialogState.UpdateAvailable);
                }
                else
                {
                    CurrentVersionLabel.Text = $"Excel Productivity Suite v{update.CurrentVersion} is currently the newest version.";
                    ShowState(DialogState.UpToDate);
                }
            }
            catch (Exception ex)
            {
                ErrorMessageText.Text = $"Failed to check for updates: {ex.Message}";
                ShowState(DialogState.Error);
            }
        }

        private async void InstallUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (_currentUpdateInfo == null || string.IsNullOrWhiteSpace(_currentUpdateInfo.DownloadUrl))
            {
                MessageBox.Show("No download package is available for this release.", "Update Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                "Microsoft Excel must close to install the update and will restart automatically.\n\nPlease save all open workbooks before continuing.\n\nDo you want to proceed with the update?",
                "Install Excel Productivity Suite Update",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes)
                return;

            ShowState(DialogState.Downloading);

            var progress = new Progress<double>(percent =>
            {
                DownloadProgressBar.Value = percent;
                DownloadPercentText.Text = $"{percent:0}%";
                DownloadDetailText.Text = percent >= 100 ? "Verifying package..." : "Downloading...";
            });

            try
            {
                string zipPath = await _updateService.DownloadUpdateAsync(_currentUpdateInfo, progress);

                DownloadStatusLabel.Text = "Launching automatic updater...";
                DownloadDetailText.Text = "Closing Excel...";

                bool launched = _updateService.LaunchInstaller(zipPath, out string error);
                if (launched)
                {
                    DialogResult = true;
                    Close();
                }
                else
                {
                    MessageBox.Show($"Failed to launch installer: {error}", "Update Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    ShowState(DialogState.UpdateAvailable);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Download failed: {ex.Message}", "Update Download Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ShowState(DialogState.UpdateAvailable);
            }
        }

        private void ViewOnGitHub_Click(object sender, RoutedEventArgs e)
        {
            string url = (_currentUpdateInfo != null && !string.IsNullOrWhiteSpace(_currentUpdateInfo.HtmlUrl))
                ? _currentUpdateInfo.HtmlUrl
                : ProductInfo.ReleasesPageUrl;

            try
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to open link in browser: {ex.Message}", "Open Link", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async void CheckAgain_Click(object sender, RoutedEventArgs e)
        {
            await CheckForUpdatesAsync();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ShowState(DialogState state)
        {
            CheckingPanel.Visibility = state == DialogState.Checking ? Visibility.Visible : Visibility.Collapsed;
            UpdateAvailablePanel.Visibility = state == DialogState.UpdateAvailable ? Visibility.Visible : Visibility.Collapsed;
            DownloadingPanel.Visibility = state == DialogState.Downloading ? Visibility.Visible : Visibility.Collapsed;
            UpToDatePanel.Visibility = state == DialogState.UpToDate ? Visibility.Visible : Visibility.Collapsed;
            ErrorPanel.Visibility = state == DialogState.Error ? Visibility.Visible : Visibility.Collapsed;

            InstallUpdateButton.Visibility = state == DialogState.UpdateAvailable ? Visibility.Visible : Visibility.Collapsed;
            CheckAgainButton.Visibility = (state == DialogState.UpToDate || state == DialogState.Error) ? Visibility.Visible : Visibility.Collapsed;
            GitHubWebButton.Visibility = state != DialogState.Checking && state != DialogState.Downloading ? Visibility.Visible : Visibility.Collapsed;
            CloseButton.IsEnabled = state != DialogState.Downloading;
        }

        private enum DialogState
        {
            Checking,
            UpdateAvailable,
            Downloading,
            UpToDate,
            Error
        }
    }
}
