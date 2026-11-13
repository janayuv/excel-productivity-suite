using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Core.Abstractions
{
    /// <summary>
    /// Service for checking, downloading, and applying product updates from GitHub.
    /// </summary>
    public interface IUpdateService
    {
        /// <summary>
        /// Checks the GitHub Releases API for newer versions of the suite.
        /// </summary>
        Task<UpdateInfo> CheckForUpdatesAsync(CancellationToken cancellationToken = default);

        /// <summary>
        /// Downloads the latest release archive and verifies its checksum.
        /// </summary>
        Task<string> DownloadUpdateAsync(UpdateInfo update, IProgress<double>? progress = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// Launches the detached installer script to safely update the add-in when Excel closes.
        /// </summary>
        bool LaunchInstaller(string downloadedZipPath, out string errorMessage);
    }
}
