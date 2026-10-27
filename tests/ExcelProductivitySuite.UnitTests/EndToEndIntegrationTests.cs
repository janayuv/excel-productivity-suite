using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Core.Services;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.Engine.Worksheets;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class EndToEndIntegrationTests
    {
        private CommandRegistry _registry = null!;
        private MockExcelContext _context = null!;

        [TestInitialize]
        public void Setup()
        {
            _registry = new CommandRegistry();
            _context = new MockExcelContext();

            // Register all suite commands
            _registry.RegisterCommand(new SuiteCommand(
                "CMD_LAUNCHER", "Tool Launcher", "Search tools", CommandCategory.General,
                ctx => ctx.ShowMessage("Launcher opened", "Launcher")
            ));

            _registry.RegisterCommand(new SuiteCommand(
                "CMD_FORMULA_TO_VALUE", "Formulas to Values", "Converts formulas to values", CommandCategory.DataCleaning,
                ctx =>
                {
                    int count = ctx.ConvertSelectionFormulasToValues();
                    ctx.ShowMessage($"Converted {count} cells.", "Formulas to Values");
                }
            ));

            _registry.RegisterCommand(new SuiteCommand(
                "CMD_TEXT_EXTRACT", "Extract Text", "Extracts substrings", CommandCategory.Text,
                ctx =>
                {
                    var matrix = ctx.GetSelectedValues();
                    if (matrix == null) return;
                    var opts = new TextExtractionOptions { Mode = ExtractionMode.BeforeDelimiter, PrimaryDelimiter = "-" };
                    var updated = (object?[,])matrix.Clone();
                    for (int r = matrix.GetLowerBound(0); r <= matrix.GetUpperBound(0); r++)
                    {
                        for (int c = matrix.GetLowerBound(1); c <= matrix.GetUpperBound(1); c++)
                        {
                            updated[r, c] = TextExtractor.Extract(matrix[r, c]?.ToString() ?? "", opts);
                        }
                    }
                    ctx.SetSelectedValues(updated);
                    ctx.ShowMessage("Extracted", "Success");
                }
            ));

            _registry.RegisterCommand(new SuiteCommand(
                "CMD_RANGE_FILL_BLANKS", "Fill Blank Cells", "Fills blanks", CommandCategory.RangeAndCells,
                ctx =>
                {
                    var matrix = ctx.GetSelectedValues();
                    if (matrix == null) return;
                    var opts = new FillBlankOptions { Direction = FillBlankDirection.Down };
                    var res = RangeFill.FillBlanks(matrix, opts, out int cnt);
                    ctx.SetSelectedValues(res);
                    ctx.ShowMessage($"Filled {cnt} blanks", "Success");
                }
            ));

            _registry.RegisterCommand(new SuiteCommand(
                "CMD_DATA_REMOVE_DUPLICATES", "Remove Duplicates", "Clears duplicate rows", CommandCategory.DataCleaning,
                ctx =>
                {
                    var matrix = ctx.GetSelectedValues();
                    if (matrix == null) return;
                    var opts = new DuplicateDetectionOptions { KeepStrategy = DuplicateKeepStrategy.KeepFirst };
                    var dups = DuplicateFinder.IdentifyDuplicateRows(matrix, opts);
                    var updated = (object?[,])matrix.Clone();
                    int rStart = matrix.GetLowerBound(0);
                    int cStart = matrix.GetLowerBound(1);
                    int cEnd = matrix.GetUpperBound(1);
                    int cleared = 0;
                    for (int r = 0; r < dups.Length; r++)
                    {
                        if (dups[r])
                        {
                            cleared++;
                            for (int c = cStart; c <= cEnd; c++) updated[rStart + r, c] = "";
                        }
                    }
                    ctx.SetSelectedValues(updated);
                    ctx.ShowMessage($"Cleared {cleared} duplicates", "Success");
                }
            ));
        }

        [TestMethod]
        public void Registry_Search_FindsCommandsByKeyword()
        {
            var results = _registry.SearchCommands("extract");
            Assert.AreEqual(1, results.Count);
            Assert.AreEqual("CMD_TEXT_EXTRACT", results[0].Id);
        }

        [TestMethod]
        public void Registry_Execute_FormulasToValues_UpdatesContext()
        {
            _registry.Execute("CMD_FORMULA_TO_VALUE", _context);
            Assert.AreEqual(1, _context.MessagesShown.Count);
            Assert.IsTrue(_context.MessagesShown[0].Contains("Converted 20 cells"));
        }

        [TestMethod]
        public void Registry_Execute_ExtractText_TransformsValues()
        {
            _context.CurrentValues = new object?[2, 1]
            {
                { "PROD-100" },
                { "ITEM-200" }
            };

            _registry.Execute("CMD_TEXT_EXTRACT", _context);

            Assert.AreEqual("PROD", _context.CurrentValues[0, 0]);
            Assert.AreEqual("ITEM", _context.CurrentValues[1, 0]);
        }

        [TestMethod]
        public void Registry_Execute_FillBlanks_FillsDownValues()
        {
            _context.CurrentValues = new object?[3, 1]
            {
                { "Category A" },
                { null },
                { null }
            };

            _registry.Execute("CMD_RANGE_FILL_BLANKS", _context);

            Assert.AreEqual("Category A", _context.CurrentValues[0, 0]);
            Assert.AreEqual("Category A", _context.CurrentValues[1, 0]);
            Assert.AreEqual("Category A", _context.CurrentValues[2, 0]);
        }

        [TestMethod]
        public void Registry_Execute_RemoveDuplicates_ClearsDuplicateRow()
        {
            _context.CurrentValues = new object?[3, 1]
            {
                { "Alpha" },
                { "Beta" },
                { "Alpha" } // Duplicate
            };

            _registry.Execute("CMD_DATA_REMOVE_DUPLICATES", _context);

            Assert.AreEqual("Alpha", _context.CurrentValues[0, 0]);
            Assert.AreEqual("Beta", _context.CurrentValues[1, 0]);
            Assert.AreEqual("", _context.CurrentValues[2, 0]);
        }
    }
}
