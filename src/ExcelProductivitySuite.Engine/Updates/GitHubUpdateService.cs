using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelProductivitySuite.Core;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Updates
{
    /// <summary>
    /// Service that checks GitHub Releases for new suite versions, downloads release zips, and applies updates.
    /// </summary>
    public class GitHubUpdateService : IUpdateService
    {
        private readonly string _endpointUrl;
        private readonly string _currentVersion;

        public GitHubUpdateService(string? endpointUrl = null, string? currentVersion = null)
        {
            _endpointUrl = endpointUrl ?? ProductInfo.ReleasesApiUrl;
            _currentVersion = currentVersion ?? ProductInfo.CurrentVersion;
        }

        public async Task<UpdateInfo> CheckForUpdatesAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                using (var handler = new HttpClientHandler { AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate })
                using (var client = new HttpClient(handler))
                {
                    client.Timeout = TimeSpan.FromSeconds(15);
                    client.DefaultRequestHeaders.Add("User-Agent", $"{ProductInfo.ProductName}-Updater/{_currentVersion}");
                    client.DefaultRequestHeaders.Add("Accept", "application/vnd.github.v3+json");

                    var response = await client.GetAsync(_endpointUrl, cancellationToken).ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        return UpdateInfo.CreateError(_currentVersion, $"GitHub returned {(int)response.StatusCode}: {response.ReasonPhrase}");
                    }

                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return GitHubReleaseParser.Parse(json, _currentVersion);
                }
            }
            catch (OperationCanceledException)
            {
                return UpdateInfo.CreateError(_currentVersion, "Update check cancelled.");
            }
            catch (Exception ex)
            {
                return UpdateInfo.CreateError(_currentVersion, $"Unable to check GitHub for updates: {ex.Message}");
            }
        }

        public async Task<string> DownloadUpdateAsync(UpdateInfo update, IProgress<double>? progress = null, CancellationToken cancellationToken = default)
        {
            if (update == null || string.IsNullOrWhiteSpace(update.DownloadUrl))
                throw new ArgumentException("No download URL provided in update info.", nameof(update));

            string updateDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelProductivitySuite", "Updates");
            if (!Directory.Exists(updateDir))
            {
                Directory.CreateDirectory(updateDir);
            }

            string fileName = string.IsNullOrWhiteSpace(update.AssetName) ? $"ExcelProductivitySuite-v{update.LatestVersion}-win-x64.zip" : update.AssetName;
            string targetPath = Path.Combine(updateDir, fileName);

            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromMinutes(5);
                client.DefaultRequestHeaders.Add("User-Agent", $"{ProductInfo.ProductName}-Updater/{_currentVersion}");

                using (var response = await client.GetAsync(update.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
                {
                    response.EnsureSuccessStatusCode();

                    long? totalBytes = response.Content.Headers.ContentLength;
                    if (!totalBytes.HasValue || totalBytes.Value <= 0)
                        totalBytes = update.AssetSizeBytes > 0 ? update.AssetSizeBytes : null;

                    using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (var fileStream = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
                    {
                        byte[] buffer = new byte[81920];
                        long totalRead = 0;
                        int bytesRead;

                        while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, bytesRead, cancellationToken).ConfigureAwait(false);
                            totalRead += bytesRead;

                            if (totalBytes.HasValue && totalBytes.Value > 0)
                            {
                                double percentage = (double)totalRead / totalBytes.Value * 100.0;
                                progress?.Report(Math.Min(100.0, percentage));
                            }
                        }
                    }
                }

                // Verify SHA-256 against release SHA256SUMS.txt if available
                if (!string.IsNullOrWhiteSpace(update.ChecksumUrl))
                {
                    try
                    {
                        string checksumsText = await client.GetStringAsync(update.ChecksumUrl).ConfigureAwait(false);
                        string computedHash = ComputeSha256(targetPath);
                        if (!string.IsNullOrEmpty(computedHash) && checksumsText.IndexOf(computedHash, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            Debug.WriteLine($"Checksum mismatch warning: {computedHash} not found in SHA256SUMS.txt");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Checksum check error: {ex.Message}");
                    }
                }
            }

            progress?.Report(100.0);
            return targetPath;
        }

        public bool LaunchInstaller(string downloadedZipPath, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!File.Exists(downloadedZipPath))
            {
                errorMessage = "Downloaded update archive not found on disk.";
                return false;
            }

            try
            {
                string updateDir = Path.GetDirectoryName(downloadedZipPath) ?? Path.GetTempPath();
                string installDir = GetInstallationDirectory();

                string scriptPath = Path.Combine(updateDir, "apply-update.cmd");
                string scriptContent = GenerateUpdaterScript(downloadedZipPath, installDir);
                File.WriteAllText(scriptPath, scriptContent, Encoding.Default);

                var startInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c \"\"{scriptPath}\"\"",
                    WorkingDirectory = updateDir,
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Normal
                };

                Process.Start(startInfo);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        public static string ComputeSha256(string filePath)
        {
            if (!File.Exists(filePath)) return string.Empty;
            using (var sha = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hash = sha.ComputeHash(stream);
                var sb = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                {
                    sb.Append(b.ToString("X2"));
                }
                return sb.ToString();
            }
        }

        public static string GetInstallationDirectory()
        {
            string localProgramsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "ExcelProductivitySuite");
            if (Directory.Exists(localProgramsDir) && File.Exists(Path.Combine(localProgramsDir, "ExcelProductivitySuite.AddIn.dll")))
            {
                return localProgramsDir;
            }

            try
            {
                string runningDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
                if (!string.IsNullOrEmpty(runningDir) && Directory.Exists(runningDir))
                {
                    return runningDir;
                }
            }
            catch { }

            return localProgramsDir;
        }

        public static string GenerateUpdaterScript(string zipPath, string targetDir)
        {
            string zipFileName = Path.GetFileName(zipPath);
            return $@"@echo off
chcp 65001 >nul
title Excel Productivity Suite - Updating...
echo ========================================================
echo   Excel Productivity Suite - Automatic Updater
echo ========================================================
echo.
echo Please save all open Excel workbooks.
echo Waiting for Microsoft Excel to close...
echo.

:wait_loop
tasklist /fi ""imagename eq EXCEL.EXE"" 2>NUL | find /i ""EXCEL.EXE"" >NUL
if ""%ERRORLEVEL%""==""0"" (
    echo Excel is currently running. Please close Excel to complete update...
    timeout /t 2 /nobreak >nul
    goto wait_loop
)

echo.
echo Excel is closed. Applying update...
echo Extracting package...

set ""ZIPFILE={zipFileName}""
set ""TARGETDIR={targetDir}""
set ""TEMPEXTRACT=%~dp0extracted_update""

if exist ""%TEMPEXTRACT%"" rd /s /q ""%TEMPEXTRACT%""
powershell -NoProfile -ExecutionPolicy Bypass -Command ""Expand-Archive -LiteralPath '%~dp0%ZIPFILE%' -DestinationPath '%TEMPEXTRACT%' -Force""

if not exist ""%TEMPEXTRACT%"" (
    echo [ERROR] Failed to extract update package.
    pause
    exit /b 1
)

if not exist ""%TARGETDIR%"" mkdir ""%TARGETDIR%""
echo Copying updated files to %TARGETDIR%...
xcopy /s /e /y /q ""%TEMPEXTRACT%\*"" ""%TARGETDIR%\"" >nul

echo Registering updated components and unblocking files...
powershell -NoProfile -ExecutionPolicy Bypass -Command ""Get-ChildItem -Path '%TARGETDIR%' -Recurse | Unblock-File -ErrorAction SilentlyContinue""
if exist ""%TARGETDIR%\install.ps1"" (
    powershell -NoProfile -ExecutionPolicy Bypass -File ""%TARGETDIR%\install.ps1"" -TargetDir ""%TARGETDIR%"" >nul 2>&1
)

echo Cleaning up temporary files...
rd /s /q ""%TEMPEXTRACT%"" >nul 2>&1
del ""%~dp0%ZIPFILE%"" >nul 2>&1

echo.
echo ========================================================
echo   Update installed successfully!
echo   Restarting Microsoft Excel...
echo ========================================================
start """" ""excel.exe""
timeout /t 2 >nul
exit /b 0
";
        }
    }
}
