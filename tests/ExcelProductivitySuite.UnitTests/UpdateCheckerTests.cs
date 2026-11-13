using System;
using System.Threading;
using ExcelProductivitySuite.Core;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Updates;
using ExcelProductivitySuite.UI.Dialogs;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class UpdateCheckerTests
    {
        private static void RunOnSta(Action action)
        {
            Exception? caught = null;
            var thread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    caught = ex;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (caught != null)
            {
                throw new Exception($"STA execution failed: {caught.Message}", caught);
            }
        }

        [TestMethod]
        public void VersionComparison_DetectsNewerVersionsAccurately()
        {
            // Newer patch
            Assert.IsTrue(GitHubReleaseParser.IsNewerVersion("5.8.0", "5.8.1"));
            Assert.IsTrue(GitHubReleaseParser.IsNewerVersion("v5.8.0", "v5.8.1"));

            // Newer minor
            Assert.IsTrue(GitHubReleaseParser.IsNewerVersion("5.8.0", "5.9.0"));
            Assert.IsTrue(GitHubReleaseParser.IsNewerVersion("v5.8.0", "v5.9.0"));

            // Newer major
            Assert.IsTrue(GitHubReleaseParser.IsNewerVersion("5.8.0", "6.0.0"));

            // Identical versions
            Assert.IsFalse(GitHubReleaseParser.IsNewerVersion("5.8.0", "5.8.0"));
            Assert.IsFalse(GitHubReleaseParser.IsNewerVersion("v5.8.0", "v5.8.0"));
            Assert.IsFalse(GitHubReleaseParser.IsNewerVersion("5.8.0", "v5.8.0"));

            // Older versions
            Assert.IsFalse(GitHubReleaseParser.IsNewerVersion("5.8.0", "5.7.0"));
            Assert.IsFalse(GitHubReleaseParser.IsNewerVersion("v5.8.0", "v5.6.0"));

            // Short forms
            Assert.IsTrue(GitHubReleaseParser.IsNewerVersion("5.8", "5.8.1"));
            Assert.IsFalse(GitHubReleaseParser.IsNewerVersion("5.8.1", "5.8"));
        }

        [TestMethod]
        public void NormalizeVersionTag_StripsVPrefixProperly()
        {
            Assert.AreEqual("5.8.0", GitHubReleaseParser.NormalizeVersionTag("v5.8.0"));
            Assert.AreEqual("5.8.0", GitHubReleaseParser.NormalizeVersionTag("V5.8.0"));
            Assert.AreEqual("5.8.0", GitHubReleaseParser.NormalizeVersionTag("5.8.0"));
            Assert.AreEqual("", GitHubReleaseParser.NormalizeVersionTag(""));
        }

        [TestMethod]
        public void GitHubReleaseParser_ParsesValidJsonPayload()
        {
            string json = @"{
                ""tag_name"": ""v5.9.0"",
                ""name"": ""Excel Productivity Suite v5.9.0"",
                ""html_url"": ""https://github.com/janayuv/excel-productivity-suite/releases/tag/v5.9.0"",
                ""published_at"": ""2026-11-20T12:00:00Z"",
                ""body"": ""## What's New\r\n- Added GitHub automatic updates\r\n- Improved performance"",
                ""assets"": [
                    {
                        ""name"": ""ExcelProductivitySuite-v5.9.0-win-x64.zip"",
                        ""size"": 590000,
                        ""browser_download_url"": ""https://github.com/janayuv/excel-productivity-suite/releases/download/v5.9.0/ExcelProductivitySuite-v5.9.0-win-x64.zip""
                    },
                    {
                        ""name"": ""SHA256SUMS.txt"",
                        ""size"": 1800,
                        ""browser_download_url"": ""https://github.com/janayuv/excel-productivity-suite/releases/download/v5.9.0/SHA256SUMS.txt""
                    }
                ]
            }";

            var update = GitHubReleaseParser.Parse(json, "5.8.0");

            Assert.IsNotNull(update);
            Assert.IsNull(update.ErrorMessage);
            Assert.IsTrue(update.IsUpdateAvailable);
            Assert.AreEqual("5.8.0", update.CurrentVersion);
            Assert.AreEqual("5.9.0", update.LatestVersion);
            Assert.AreEqual("Excel Productivity Suite v5.9.0", update.ReleaseName);
            Assert.IsTrue(update.ReleaseNotes.Contains("Added GitHub automatic updates"));
            Assert.AreEqual("https://github.com/janayuv/excel-productivity-suite/releases/download/v5.9.0/ExcelProductivitySuite-v5.9.0-win-x64.zip", update.DownloadUrl);
            Assert.AreEqual("https://github.com/janayuv/excel-productivity-suite/releases/download/v5.9.0/SHA256SUMS.txt", update.ChecksumUrl);
            Assert.AreEqual(590000, update.AssetSizeBytes);
            Assert.IsNotNull(update.PublishedAt);
        }

        [TestMethod]
        public void GitHubReleaseParser_HandlesSameVersionAsNotAvailable()
        {
            string json = @"{
                ""tag_name"": ""v5.8.0"",
                ""name"": ""Excel Productivity Suite v5.8.0"",
                ""html_url"": ""https://github.com/janayuv/excel-productivity-suite/releases/tag/v5.8.0"",
                ""body"": ""Release v5.8.0"",
                ""assets"": []
            }";

            var update = GitHubReleaseParser.Parse(json, "5.8.0");

            Assert.IsNotNull(update);
            Assert.IsFalse(update.IsUpdateAvailable, "Same version should not report an update available");
            Assert.AreEqual("5.8.0", update.LatestVersion);
        }

        [TestMethod]
        public void GitHubReleaseParser_HandlesEmptyJsonGracefully()
        {
            var update = GitHubReleaseParser.Parse("", "5.8.0");

            Assert.IsNotNull(update);
            Assert.IsFalse(update.IsUpdateAvailable);
            Assert.IsFalse(string.IsNullOrEmpty(update.ErrorMessage));
        }

        [TestMethod]
        public void GitHubUpdateService_GenerateUpdaterScript_ContainsCriticalCommands()
        {
            string script = GitHubUpdateService.GenerateUpdaterScript(@"C:\Temp\Update.zip", @"C:\Program Files\App");

            StringAssert.Contains(script, "EXCEL.EXE");
            StringAssert.Contains(script, "Expand-Archive");
            StringAssert.Contains(script, "xcopy");
            StringAssert.Contains(script, "excel.exe");
        }

        [TestMethod]
        public void CheckForUpdatesDialog_InstantiatesWithoutExceptions()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                var dlg = new CheckForUpdatesDialog(mock);

                Assert.IsNotNull(dlg);
                Assert.IsNotNull(dlg.CheckingStatePanel);
                Assert.IsNotNull(dlg.UpdateAvailableStatePanel);
                Assert.IsNotNull(dlg.DownloadingStatePanel);
                Assert.IsNotNull(dlg.UpToDateStatePanel);
                Assert.IsNotNull(dlg.ErrorStatePanel);
            });
        }
    }
}
