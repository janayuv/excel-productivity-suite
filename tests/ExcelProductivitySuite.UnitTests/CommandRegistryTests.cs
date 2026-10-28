using System;
using System.Linq;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class CommandRegistryTests
    {
        private CommandRegistry _registry = null!;
        private MockExcelContext _context = null!;

        [TestInitialize]
        public void Setup()
        {
            _registry = new CommandRegistry();
            _context = new MockExcelContext();

            // Populate sample commands across different modules and statuses
            _registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_EXTRACT",
                title: "Extract Text",
                description: "Extract substrings before, after, or between delimiters.",
                category: CommandCategory.TextTools,
                executeAction: ctx => ctx.ShowMessage("Extracted", "Text"),
                keywords: new[] { "substring", "delimiter", "parse" },
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "A",
                roadmapBatch: "Active"
            ));

            _registry.RegisterCommand(new SuiteCommand(
                id: "CMD_TEXT_REVERSE",
                title: "Reverse Text",
                description: "Reverse cell text by character or word order.",
                category: CommandCategory.TextTools,
                executeAction: ctx => ctx.ShowMessage("Reversed", "Text"),
                keywords: new[] { "reverse", "flip", "backwards" },
                status: FeatureStatus.Planned,
                moduleCode: "A",
                roadmapBatch: "Batch 2"
            ));

            _registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_TRANSPOSE",
                title: "Transpose Range",
                description: "Transpose rows to columns and columns to rows.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx => ctx.ShowMessage("Transposed", "Range"),
                keywords: new[] { "transpose", "rows to columns", "flip matrix" },
                status: FeatureStatus.Planned,
                moduleCode: "B",
                roadmapBatch: "Batch 3"
            ));

            _registry.RegisterCommand(new SuiteCommand(
                id: "CMD_RANGE_FILL_BLANKS",
                title: "Fill Blank Cells",
                description: "Fill empty cells with values from above or constant.",
                category: CommandCategory.RangeAndCells,
                executeAction: ctx => ctx.ShowMessage("Filled", "Range"),
                keywords: new[] { "fill blanks", "interpolate" },
                status: FeatureStatus.ImplementedAndVerified,
                moduleCode: "B",
                roadmapBatch: "Active"
            ));

            _registry.RegisterCommand(new SuiteCommand(
                id: "CMD_SHEET_COMBINE",
                title: "Combine Worksheets",
                description: "Combine multiple worksheets into one master sheet.",
                category: CommandCategory.WorksheetAndWorkbook,
                executeAction: ctx => ctx.ShowMessage("Combined", "Sheets"),
                keywords: new[] { "combine sheets", "merge worksheets" },
                status: FeatureStatus.Planned,
                moduleCode: "D",
                roadmapBatch: "Batch 5"
            ));
        }

        [TestMethod]
        public void RegisterAndRetrieveCommand_ReturnsCorrectCommand()
        {
            var cmd = _registry.GetCommand("CMD_TEXT_EXTRACT");
            Assert.IsNotNull(cmd);
            Assert.AreEqual("Extract Text", cmd.Title);
            Assert.AreEqual(CommandCategory.TextTools, cmd.Category);
            Assert.AreEqual(FeatureStatus.ImplementedAndVerified, cmd.Status);
            Assert.IsTrue(cmd.IsImplemented);
            Assert.AreEqual("Active", cmd.StatusDisplay);
            Assert.AreEqual("A", cmd.ModuleCode);
        }

        [TestMethod]
        public void SearchCommands_ByKeyword_ReturnsRankedMatches()
        {
            var results = _registry.SearchCommands("delimiter");
            Assert.IsTrue(results.Count > 0);
            Assert.AreEqual("CMD_TEXT_EXTRACT", results[0].Id);
        }

        [TestMethod]
        public void SearchCommands_ByModuleCode_MatchesModuleTools()
        {
            var results = _registry.SearchCommands("A");
            Assert.IsTrue(results.Any(c => c.ModuleCode == "A"));
        }

        [TestMethod]
        public void SearchCommands_FilterByStatus_ActiveOnly()
        {
            var active = _registry.SearchCommands("", null, FeatureStatus.ImplementedAndVerified);
            Assert.AreEqual(2, active.Count);
            Assert.IsTrue(active.All(c => c.IsImplemented));
        }

        [TestMethod]
        public void SearchCommands_FilterByStatus_PlannedOnly()
        {
            var planned = _registry.SearchCommands("", null, FeatureStatus.Planned);
            Assert.AreEqual(3, planned.Count);
            Assert.IsTrue(planned.All(c => !c.IsImplemented));
        }

        [TestMethod]
        public void SearchCommands_FilterByCategory()
        {
            var textTools = _registry.SearchCommands("", CommandCategory.TextTools);
            Assert.AreEqual(2, textTools.Count);
            Assert.IsTrue(textTools.All(c => c.Category == CommandCategory.TextTools));

            var rangeTools = _registry.SearchCommands("", CommandCategory.RangeAndCells);
            Assert.AreEqual(2, rangeTools.Count);
        }

        [TestMethod]
        public void BackwardCompatibility_EnumAliasesWorkSeamlessly()
        {
            // CommandCategory.Text is an alias for TextTools
            Assert.AreEqual((int)CommandCategory.TextTools, (int)CommandCategory.Text);
            Assert.AreEqual((int)CommandCategory.DataAndFiltering, (int)CommandCategory.DataCleaning);
            Assert.AreEqual((int)CommandCategory.WorksheetAndWorkbook, (int)CommandCategory.Worksheet);
            Assert.AreEqual((int)CommandCategory.FormulasAndCalculation, (int)CommandCategory.Formulas);
        }

        [TestMethod]
        public void ExecuteCommand_InvokesActionSuccessfully()
        {
            bool executed = false;
            _registry.RegisterCommand(new SuiteCommand(
                "CMD_TEST_EXEC", "Test Tool", "Testing execution", CommandCategory.General,
                ctx => { executed = true; }
            ));

            _registry.Execute("CMD_TEST_EXEC", _context);
            Assert.IsTrue(executed);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ExecuteCommand_UnknownId_ThrowsInvalidOperationException()
        {
            _registry.Execute("CMD_NON_EXISTENT", _context);
        }
    }
}
