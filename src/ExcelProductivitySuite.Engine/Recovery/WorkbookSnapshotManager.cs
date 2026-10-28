using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelProductivitySuite.Engine.Recovery
{
    public class SnapshotFileInfo
    {
        public string FileName { get; }
        public string FilePath { get; }
        public string WorkbookName { get; }
        public DateTime CreatedTime { get; }
        public long FileSizeBytes { get; }
        public string FormattedSize { get; }
        public string Tag { get; }

        public SnapshotFileInfo(string filePath, string workbookName, DateTime createdTime, long fileSizeBytes, string tag)
        {
            FilePath = filePath;
            FileName = Path.GetFileName(filePath);
            WorkbookName = workbookName;
            CreatedTime = createdTime;
            FileSizeBytes = fileSizeBytes;
            Tag = tag;
            FormattedSize = FormatFileSize(fileSizeBytes);
        }

        private static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
            return $"{bytes / (1024.0 * 1024.0):F2} MB";
        }
    }

    public static class WorkbookSnapshotManager
    {
        public static string GetDefaultSnapshotDirectory()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string snapshotDir = Path.Combine(localAppData, "ExcelProductivitySuite", "Snapshots");
            if (!Directory.Exists(snapshotDir))
            {
                Directory.CreateDirectory(snapshotDir);
            }
            return snapshotDir;
        }

        public static string GenerateSnapshotFilePath(string directory, string workbookName, string? customTag = null)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                directory = GetDefaultSnapshotDirectory();
            }

            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string rawName = Path.GetFileNameWithoutExtension(workbookName) ?? "Workbook";
            string cleanWbName = SanitizeFileName(rawName);
            if (string.IsNullOrWhiteSpace(cleanWbName))
            {
                cleanWbName = "Workbook";
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string tagSuffix = "";
            if (!string.IsNullOrWhiteSpace(customTag))
            {
                string cleanTag = SanitizeFileName(customTag!);
                if (!string.IsNullOrWhiteSpace(cleanTag))
                {
                    tagSuffix = $"_{cleanTag}";
                }
            }

            string ext = Path.GetExtension(workbookName);
            if (string.IsNullOrWhiteSpace(ext) || (!ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) && !ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase)))
            {
                ext = ".xlsx";
            }

            string fileName = $"{cleanWbName}_{timestamp}{tagSuffix}{ext}";
            return Path.Combine(directory, fileName);
        }

        public static IReadOnlyList<SnapshotFileInfo> GetSnapshots(string? directory = null, string? workbookFilter = null)
        {
            string dir = string.IsNullOrWhiteSpace(directory) ? GetDefaultSnapshotDirectory() : directory!;
            var results = new List<SnapshotFileInfo>();

            if (!Directory.Exists(dir))
            {
                return results;
            }

            var dirInfo = new DirectoryInfo(dir);
            var files = dirInfo.GetFiles("*.xls*").OrderByDescending(f => f.CreationTime);

            string? filterClean = string.IsNullOrWhiteSpace(workbookFilter) ? null : Path.GetFileNameWithoutExtension(workbookFilter);

            foreach (var fi in files)
            {
                string nameWithoutExt = Path.GetFileNameWithoutExtension(fi.Name);
                // Parse filename parts: e.g. "Sales_20261028_153000_BeforePivot"
                string[] parts = nameWithoutExt.Split('_');
                string wbName = parts[0];
                string tag = parts.Length > 3 ? string.Join("_", parts.Skip(3)) : (parts.Length == 3 && !parts[2].All(char.IsDigit) ? parts[2] : "Snapshot");

                if (!string.IsNullOrWhiteSpace(filterClean))
                {
                    if (!wbName.Equals(filterClean, StringComparison.OrdinalIgnoreCase) &&
                        !fi.Name.StartsWith(filterClean, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }

                results.Add(new SnapshotFileInfo(fi.FullName, wbName, fi.CreationTime, fi.Length, tag));
            }

            return results;
        }

        public static bool DeleteSnapshot(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return false;
            try
            {
                File.Delete(filePath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static int CleanOldSnapshots(string? directory = null, int maxSnapshots = 30)
        {
            string dir = string.IsNullOrWhiteSpace(directory) ? GetDefaultSnapshotDirectory() : directory!;
            if (!Directory.Exists(dir)) return 0;

            int deleted = 0;
            var dirInfo = new DirectoryInfo(dir);
            var files = dirInfo.GetFiles("*.xls*").OrderByDescending(f => f.CreationTime).ToList();

            if (files.Count > maxSnapshots)
            {
                var toDelete = files.Skip(maxSnapshots);
                foreach (var f in toDelete)
                {
                    try
                    {
                        f.Delete();
                        deleted++;
                    }
                    catch { }
                }
            }

            return deleted;
        }

        private static string SanitizeFileName(string name)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            var cleaned = new char[name.Length];
            int idx = 0;
            for (int i = 0; i < name.Length; i++)
            {
                char c = name[i];
                if (invalidChars.Contains(c) || c == ' ')
                {
                    cleaned[idx++] = '_';
                }
                else
                {
                    cleaned[idx++] = c;
                }
            }
            return new string(cleaned, 0, idx).Trim('_');
        }
    }
}
