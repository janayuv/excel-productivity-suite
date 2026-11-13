using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class RibbonTests
    {
        private static readonly XNamespace CustomUiNs = "http://schemas.microsoft.com/office/2009/07/customui";

        [TestMethod]
        public void RibbonXml_ExcelProductivityTab_IsLastAndHasNoInsertBeforeMso()
        {
            var addInAsm = LoadAddInAssembly();
            object ribbon = CreateRibbonInstance(addInAsm);
            string xml = InvokeGetCustomUI(ribbon);

            var doc = XDocument.Parse(xml);
            var root = doc.Root;
            Assert.IsNotNull(root);
            Assert.AreEqual("LoadImage", root.Attribute("loadImage")?.Value);

            var tabs = root.Descendants(CustomUiNs + "tab").ToList();
            Assert.AreEqual(2, tabs.Count);
            Assert.AreEqual("tabWorkbookAndSheets", tabs[0].Attribute("id")?.Value);
            Assert.AreEqual("tabExcelProductivity", tabs[1].Attribute("id")?.Value);
            Assert.IsTrue(
                tabs.All(t => t.Attribute("insertBeforeMso") == null),
                "No custom tab should declare insertBeforeMso so Excel Productivity remains in the last position.");
        }

        [TestMethod]
        public void RibbonXml_AllGroupsAndButtons_HaveValidEmbedded32bppIcons()
        {
            var addInAsm = LoadAddInAssembly();
            object ribbon = CreateRibbonInstance(addInAsm);
            string xml = InvokeGetCustomUI(ribbon);

            var doc = XDocument.Parse(xml);
            Assert.IsNotNull(doc.Root);
            Assert.IsFalse(
                doc.Descendants().Any(e => e.Attribute("imageMso") != null),
                "All built-in imageMso references must be replaced with custom image attributes.");

            var groups = doc.Root.Descendants(CustomUiNs + "group").ToList();
            var buttons = doc.Root.Descendants(CustomUiNs + "button").ToList();
            Assert.AreEqual(15, groups.Count);
            Assert.AreEqual(118, buttons.Count);

            var expectedIcons = new List<(string Id, string ImageAttr, int ExpectedPx)>();
            foreach (var g in groups)
            {
                string gid = g.Attribute("id")?.Value ?? "(unnamed-group)";
                string? img = g.Attribute("image")?.Value;
                Assert.IsFalse(string.IsNullOrWhiteSpace(img), $"Group '{gid}' is missing image attribute.");
                StringAssert.StartsWith(img, "large/", $"Group '{gid}' must use a 32x32 large/ icon.");
                expectedIcons.Add((gid, img!, 32));
            }

            foreach (var b in buttons)
            {
                string bid = b.Attribute("id")?.Value ?? "(unnamed-button)";
                string? size = b.Attribute("size")?.Value;
                string? img = b.Attribute("image")?.Value;
                Assert.IsFalse(string.IsNullOrWhiteSpace(img), $"Button '{bid}' is missing image attribute.");

                int expectedPx = string.Equals(size, "large", StringComparison.Ordinal) ? 32 : 16;
                string expectedPrefix = expectedPx == 32 ? "large/" : "small/";
                StringAssert.StartsWith(img, expectedPrefix, $"Button '{bid}' (size='{size}') must use '{expectedPrefix}'.");
                expectedIcons.Add((bid, img!, expectedPx));
            }

            RunOnStaThread(() =>
            {
                MethodInfo? resolveNameMethod = ribbon.GetType().GetMethod(
                    "ResolveIconResourceName",
                    BindingFlags.Static | BindingFlags.NonPublic);
                MethodInfo? loadImageMethod = ribbon.GetType().GetMethod(
                    "LoadImage",
                    BindingFlags.Instance | BindingFlags.Public);
                Assert.IsNotNull(resolveNameMethod);
                Assert.IsNotNull(loadImageMethod);

                foreach (var (id, imageAttr, expectedPx) in expectedIcons)
                {
                    string resName = (string)resolveNameMethod!.Invoke(null, new object[] { imageAttr })!;
                    using (var stream = addInAsm.GetManifestResourceStream(resName))
                    {
                        Assert.IsNotNull(stream, $"Embedded icon resource '{resName}' missing for '{id}'.");
                        AssertPngRgbaDimensions(stream!, expectedPx, expectedPx, resName);
                    }

                    object? pic = loadImageMethod!.Invoke(ribbon, new object[] { imageAttr });
                    Assert.IsNotNull(pic, $"LoadImage returned null for '{imageAttr}' (control '{id}').");

                    short picType = (short)pic!.GetType().InvokeMember("Type", BindingFlags.GetProperty, null, pic, null)!;
                    int handleVal = (int)pic.GetType().InvokeMember("Handle", BindingFlags.GetProperty, null, pic, null)!;
                    Assert.AreEqual(1, picType, $"Expected PICTYPE_BITMAP for '{imageAttr}'.");
                    Assert.AreNotEqual(0, handleVal, $"Expected non-zero HBITMAP handle for '{imageAttr}'.");

                    var bmpInfo = new Win32Bitmap();
                    int bytesRead = GetObjectBitmap(new IntPtr(handleVal), Marshal.SizeOf(typeof(Win32Bitmap)), ref bmpInfo);
                    Assert.IsTrue(bytesRead > 0, $"GetObject failed for '{imageAttr}'.");
                    Assert.AreEqual(expectedPx, bmpInfo.bmWidth, $"Width mismatch for '{imageAttr}'.");
                    Assert.AreEqual(expectedPx, bmpInfo.bmHeight, $"Height mismatch for '{imageAttr}'.");
                    Assert.AreEqual((ushort)32, bmpInfo.bmBitsPixel, $"Expected 32bpp DIB for '{imageAttr}'.");
                }
            });
        }

        private static void AssertPngRgbaDimensions(Stream stream, int expectedWidth, int expectedHeight, string resourceName)
        {
            byte[] header = new byte[26];
            int read = stream.Read(header, 0, header.Length);
            Assert.AreEqual(header.Length, read, $"Truncated PNG header in '{resourceName}'.");
            Assert.AreEqual(0x89, header[0]);
            Assert.AreEqual((byte)'P', header[1]);
            Assert.AreEqual((byte)'N', header[2]);
            Assert.AreEqual((byte)'G', header[3]);

            int width = (header[16] << 24) | (header[17] << 16) | (header[18] << 8) | header[19];
            int height = (header[20] << 24) | (header[21] << 16) | (header[22] << 8) | header[23];
            byte bitDepth = header[24];
            byte colorType = header[25];

            Assert.AreEqual(expectedWidth, width, $"PNG width mismatch in '{resourceName}'.");
            Assert.AreEqual(expectedHeight, height, $"PNG height mismatch in '{resourceName}'.");
            Assert.AreEqual((byte)8, bitDepth, $"Expected 8-bit PNG in '{resourceName}'.");
            Assert.AreEqual((byte)6, colorType, $"Expected RGBA (type 6) PNG with alpha in '{resourceName}'.");
        }

        private static Assembly LoadAddInAssembly()
        {
            string baseDir = AppContext.BaseDirectory;
            string config = baseDir.IndexOf("Release", StringComparison.OrdinalIgnoreCase) >= 0 ? "Release" : "Debug";
            string repoRoot = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", ".."));
            string candidate = Path.Combine(repoRoot, "src", "ExcelProductivitySuite.AddIn", "bin", config, "ExcelProductivitySuite.AddIn.dll");
            if (!File.Exists(candidate))
            {
                string fallback = config == "Release" ? "Debug" : "Release";
                candidate = Path.Combine(repoRoot, "src", "ExcelProductivitySuite.AddIn", "bin", fallback, "ExcelProductivitySuite.AddIn.dll");
            }

            Assert.IsTrue(File.Exists(candidate), $"AddIn assembly not found at '{candidate}'. Build the solution before running RibbonTests.");
            return Assembly.LoadFrom(candidate);
        }

        private static object CreateRibbonInstance(Assembly addInAsm)
        {
            var ribbonType = addInAsm.GetType("ExcelProductivitySuite.AddIn.Ribbon.ExcelProductivityRibbon", throwOnError: true)!;
            var instance = Activator.CreateInstance(ribbonType, new StubCommandRegistry(), new MockExcelContext());
            Assert.IsNotNull(instance);
            return instance!;
        }

        private static string InvokeGetCustomUI(object ribbon)
        {
            var method = ribbon.GetType().GetMethod("GetCustomUI", BindingFlags.Instance | BindingFlags.Public);
            Assert.IsNotNull(method);
            return (string)method!.Invoke(ribbon, new object[] { "Microsoft.Excel.Workbook" })!;
        }

        private static void RunOnStaThread(Action action)
        {
            Exception? captured = null;
            var thread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    captured = ex;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (captured != null)
                throw new AssertFailedException($"STA test execution failed: {captured}", captured);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Win32Bitmap
        {
            public int bmType;
            public int bmWidth;
            public int bmHeight;
            public int bmWidthBytes;
            public ushort bmPlanes;
            public ushort bmBitsPixel;
            public IntPtr bmBits;
        }

        [DllImport("gdi32.dll", EntryPoint = "GetObjectW")]
        private static extern int GetObjectBitmap(IntPtr hObject, int cbBuffer, ref Win32Bitmap lpvObject);

        private sealed class StubCommandRegistry : ICommandRegistry
        {
            public void RegisterCommand(SuiteCommand command) { }
            public SuiteCommand? GetCommand(string id) => null;
            public IReadOnlyList<SuiteCommand> GetAllCommands() => Array.Empty<SuiteCommand>();
            public IReadOnlyList<SuiteCommand> SearchCommands(string query, CommandCategory? categoryFilter = null) => Array.Empty<SuiteCommand>();
            public IReadOnlyList<SuiteCommand> SearchCommands(string query, CommandCategory? categoryFilter, FeatureStatus? statusFilter) => Array.Empty<SuiteCommand>();
            public void Execute(string commandId, IExcelContext context) { }
        }
    }
}
