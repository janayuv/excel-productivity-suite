using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelProductivitySuite.Engine.Files
{
    public class FolderScanOptions
    {
        public bool IncludeSubfolders { get; set; } = true;
        public string SearchPattern { get; set; } = "*.*";
        public bool GenerateHyperlinks { get; set; } = true;
    }

    public class FolderScanResult
    {
        public object?[,] TableData { get; }
        public int FileCount { get; }
        public long TotalBytes { get; }

        public FolderScanResult(object?[,] tableData, int count, long bytes)
        {
            TableData = tableData;
            FileCount = count;
            TotalBytes = bytes;
        }
    }

    /// <summary>
    /// Engine for scanning directory folders and building a structured Excel inventory table of files.
    /// </summary>
    public static class FolderInventoryScanner
    {
        public static FolderScanResult ScanFolder(string folderPath, FolderScanOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(folderPath)) throw new ArgumentNullException(nameof(folderPath));
            if (!Directory.Exists(folderPath)) throw new DirectoryNotFoundException($"Folder not found: {folderPath}");

            options ??= new FolderScanOptions();
            var searchOption = options.IncludeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            var patterns = options.SearchPattern.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (patterns.Length == 0) patterns = new[] { "*.*" };

            var files = new List<FileInfo>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var p in patterns)
            {
                try
                {
                    var di = new DirectoryInfo(folderPath);
                    var matchedFiles = di.GetFiles(p.Trim(), searchOption);
                    foreach (var f in matchedFiles)
                    {
                        if (seen.Add(f.FullName))
                        {
                            files.Add(f);
                        }
                    }
                }
                catch
                {
                    // Continue scanning other patterns if one throws permissions error
                }
            }

            int rows = files.Count + 1; // +1 for header
            int cols = 8;
            object?[,] matrix = new object?[rows, cols];

            // Headers
            matrix[0, 0] = "File Name";
            matrix[0, 1] = "Folder";
            matrix[0, 2] = "Size (Formatted)";
            matrix[0, 3] = "Size (Bytes)";
            matrix[0, 4] = "Extension";
            matrix[0, 5] = "Date Modified";
            matrix[0, 6] = "Date Created";
            matrix[0, 7] = "Link";

            long totalBytes = 0;

            for (int i = 0; i < files.Count; i++)
            {
                var f = files[i];
                int r = i + 1;
                long len = 0;
                try { len = f.Length; } catch { }
                totalBytes += len;

                matrix[r, 0] = f.Name;
                matrix[r, 1] = f.DirectoryName ?? string.Empty;
                matrix[r, 2] = FormatBytes(len);
                matrix[r, 3] = len;
                matrix[r, 4] = f.Extension.ToLowerInvariant();
                matrix[r, 5] = f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss");
                matrix[r, 6] = f.CreationTime.ToString("yyyy-MM-dd HH:mm:ss");

                if (options.GenerateHyperlinks)
                {
                    string safePath = f.FullName.Replace("\"", "\"\"");
                    string safeName = f.Name.Replace("\"", "\"\"");
                    matrix[r, 7] = $"=HYPERLINK(\"{safePath}\", \"{safeName}\")";
                }
                else
                {
                    matrix[r, 7] = f.FullName;
                }
            }

            return new FolderScanResult(matrix, files.Count, totalBytes);
        }

        public static string FormatBytes(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{(bytes / 1024.0):F1} KB";
            if (bytes < 1024 * 1024 * 1024) return $"{(bytes / (1024.0 * 1024.0)):F2} MB";
            return $"{(bytes / (1024.0 * 1024.0 * 1024.0)):F2} GB";
        }
    }
}
