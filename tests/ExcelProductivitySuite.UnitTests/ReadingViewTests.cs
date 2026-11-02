using System;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Navigation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class ReadingViewTests
    {
        [TestMethod]
        public void ReadingViewSettings_Defaults_AreValid()
        {
            var s = ReadingViewSettings.CreateDefault();

            Assert.IsFalse(s.IsEnabled);
            Assert.AreEqual(ReadingHighlightShape.RowAndColumn, s.Shape);
            Assert.AreEqual(ReadingHighlightScope.EntireRowColumn, s.Scope);
            Assert.AreEqual("#38BDF8", s.ColorHex);
            Assert.AreEqual(0.75, s.Transparency);
            Assert.IsTrue(s.BorderEnabled);
            Assert.AreEqual("#0284C7", s.BorderColorHex);

            // Hex to RGB
            int rgb = s.ColorRgb;
            Assert.AreNotEqual(0, rgb);

            // Clone
            var clone = s.Clone();
            Assert.AreEqual(s.ColorHex, clone.ColorHex);
            Assert.AreEqual(s.Transparency, clone.Transparency);
            clone.ColorHex = "#10B981";
            Assert.AreNotEqual(s.ColorHex, clone.ColorHex);
        }

        [TestMethod]
        public void ReadingLayoutEngine_CalculateGeometry_RowAndColumnCrosshair()
        {
            var s = new ReadingViewSettings
            {
                Shape = ReadingHighlightShape.RowAndColumn,
                Scope = ReadingHighlightScope.EntireRowColumn
            };

            double cellLeft = 120;
            double cellTop = 240;
            double cellWidth = 60;
            double cellHeight = 22;

            var geo = ReadingLayoutEngine.CalculateGeometry(
                cellLeft, cellTop, cellWidth, cellHeight,
                maxLeft: 0, maxTop: 0, maxWidth: 800, maxHeight: 1500,
                s);

            Assert.IsTrue(geo.HasRow);
            Assert.IsTrue(geo.HasColumn);
            Assert.IsTrue(geo.HasCell);

            // Row follows cell's Y position and height
            Assert.AreEqual(cellTop, geo.RowTop);
            Assert.AreEqual(cellHeight, geo.RowHeight);
            Assert.IsTrue(geo.RowWidth >= 4000);

            // Column follows cell's X position and width
            Assert.AreEqual(cellLeft, geo.ColLeft);
            Assert.AreEqual(cellWidth, geo.ColWidth);
            Assert.IsTrue(geo.ColHeight >= 40000);

            // Cell bounds
            Assert.AreEqual(cellLeft, geo.CellLeft);
            Assert.AreEqual(cellTop, geo.CellTop);
        }

        [TestMethod]
        public void ReadingLayoutEngine_CalculateGeometry_RowOnlyAndColOnly()
        {
            var sRow = new ReadingViewSettings { Shape = ReadingHighlightShape.RowOnly };
            var geoRow = ReadingLayoutEngine.CalculateGeometry(100, 200, 50, 20, 0, 0, 500, 500, sRow);
            Assert.IsTrue(geoRow.HasRow);
            Assert.IsFalse(geoRow.HasColumn);

            var sCol = new ReadingViewSettings { Shape = ReadingHighlightShape.ColumnOnly };
            var geoCol = ReadingLayoutEngine.CalculateGeometry(100, 200, 50, 20, 0, 0, 500, 500, sCol);
            Assert.IsFalse(geoCol.HasRow);
            Assert.IsTrue(geoCol.HasColumn);
        }

        [TestMethod]
        public void ReadingLayoutEngine_CalculateGeometry_UsedRangeScope()
        {
            var s = new ReadingViewSettings
            {
                Shape = ReadingHighlightShape.RowAndColumn,
                Scope = ReadingHighlightScope.UsedRangeOnly
            };

            double maxLeft = 50;
            double maxTop = 40;
            double maxWidth = 600;
            double maxHeight = 800;

            var geo = ReadingLayoutEngine.CalculateGeometry(
                120, 240, 60, 22,
                maxLeft, maxTop, maxWidth, maxHeight,
                s);

            Assert.AreEqual(maxLeft, geo.RowLeft);
            Assert.AreEqual(maxWidth, geo.RowWidth);

            Assert.AreEqual(maxTop, geo.ColTop);
            Assert.AreEqual(maxHeight, geo.ColHeight);
        }

        [TestMethod]
        public void ReadingLayoutEngine_SettingsSerialization_PreservesAllFields()
        {
            var orig = new ReadingViewSettings
            {
                IsEnabled = true,
                Shape = ReadingHighlightShape.ColumnOnly,
                Scope = ReadingHighlightScope.UsedRangeOnly,
                ColorHex = "#F59E0B",
                Transparency = 0.60,
                BorderEnabled = false,
                BorderColorHex = "#D97706"
            };

            string json = ReadingLayoutEngine.SerializeSettings(orig);
            var deserialized = ReadingLayoutEngine.DeserializeSettings(json);

            Assert.IsTrue(deserialized.IsEnabled);
            Assert.AreEqual(ReadingHighlightShape.ColumnOnly, deserialized.Shape);
            Assert.AreEqual(ReadingHighlightScope.UsedRangeOnly, deserialized.Scope);
            Assert.AreEqual("#F59E0B", deserialized.ColorHex);
            Assert.AreEqual(0.60, deserialized.Transparency, 0.01);
            Assert.IsFalse(deserialized.BorderEnabled);
            Assert.AreEqual("#D97706", deserialized.BorderColorHex);
        }

        [TestMethod]
        public void MockExcelContext_ReadingView_StateTransitions()
        {
            var ctx = new MockExcelContext();

            Assert.IsFalse(ctx.IsReadingViewActive);

            ctx.EnableReadingView();
            Assert.IsTrue(ctx.IsReadingViewActive);
            Assert.IsTrue(ctx.ReadingViewSettings.IsEnabled);
            Assert.IsTrue(ctx.ReadingViewHighlightUpdateCount >= 1);

            ctx.ToggleReadingView();
            Assert.IsFalse(ctx.IsReadingViewActive);

            ctx.ToggleReadingView();
            Assert.IsTrue(ctx.IsReadingViewActive);

            ctx.DisableReadingView();
            Assert.IsFalse(ctx.IsReadingViewActive);
        }

        [TestMethod]
        public void MockExcelContext_GridFocus_AliasesWorkEqually()
        {
            var ctx = new MockExcelContext();

            Assert.IsFalse(ctx.IsGridFocusActive);

            ctx.EnableGridFocus();
            Assert.IsTrue(ctx.IsGridFocusActive);
            Assert.IsTrue(ctx.IsReadingViewActive);

            ctx.ToggleGridFocus();
            Assert.IsFalse(ctx.IsGridFocusActive);
            Assert.IsFalse(ctx.IsReadingViewActive);

            ctx.ToggleGridFocus();
            Assert.IsTrue(ctx.IsGridFocusActive);

            ctx.DisableGridFocus();
            Assert.IsFalse(ctx.IsGridFocusActive);
        }
    }
}
