using System;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Represents the result of a GitHub release update check.
    /// </summary>
    public class UpdateInfo
    {
        public string CurrentVersion { get; set; } = string.Empty;
        public string LatestVersion { get; set; } = string.Empty;
        public string ReleaseName { get; set; } = string.Empty;
        public string ReleaseNotes { get; set; } = string.Empty;
        public string HtmlUrl { get; set; } = string.Empty;
        public string? DownloadUrl { get; set; }
        public string? ChecksumUrl { get; set; }
        public long AssetSizeBytes { get; set; }
        public string AssetName { get; set; } = string.Empty;
        public DateTime? PublishedAt { get; set; }
        public bool IsUpdateAvailable { get; set; }
        public string? ErrorMessage { get; set; }

        public static UpdateInfo CreateError(string currentVersion, string errorMessage)
        {
            return new UpdateInfo
            {
                CurrentVersion = currentVersion,
                ErrorMessage = errorMessage,
                IsUpdateAvailable = false
            };
        }
    }
}
