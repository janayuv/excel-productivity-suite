using System;
using System.Linq;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Navigation;
using ExcelProductivitySuite.UI.ViewModels;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class WorkbookSheetNavigatorTests
    {
        [TestMethod]
        public void SheetNameValidator_ValidatesStandardExcelRules()
        {
            var existing = new[] { "Sheet1", "Summary", "Data" };

            // Empty or whitespace
            Assert.IsFalse(SheetNameValidator.Validate("", existing, out string err1));
            Assert.IsTrue(err1.Contains("blank"));

            Assert.IsFalse(SheetNameValidator.Validate("   ", existing, out _));

            // Exceeds 31 chars
            string longName = new string('A', 32);
            Assert.IsFalse(SheetNameValidator.Validate(longName, existing, out string err2));
            Assert.IsTrue(err2.Contains("31"));

            // Illegal characters
            Assert.IsFalse(SheetNameValidator.Validate("Sales/2026", existing, out string err3));
            Assert.IsTrue(err3.Contains("characters"));
            Assert.IsFalse(SheetNameValidator.Validate("Report:Q1", existing, out _));
            Assert.IsFalse(SheetNameValidator.Validate("Data[Old]", existing, out _));
            Assert.IsFalse(SheetNameValidator.Validate("Target*", existing, out _));
            Assert.IsFalse(SheetNameValidator.Validate("Test?", existing, out _));
            Assert.IsFalse(SheetNameValidator.Validate("Back\\Slash", existing, out _));

            // Single quotes at start or end
            Assert.IsFalse(SheetNameValidator.Validate("'Quarterly", existing, out _));
            Assert.IsFalse(SheetNameValidator.Validate("Quarterly'", existing, out _));

            // Reserved name History
            Assert.IsFalse(SheetNameValidator.Validate("History", existing, out _));
            Assert.IsFalse(SheetNameValidator.Validate("history", existing, out _));

            // Duplicate name
            Assert.IsFalse(SheetNameValidator.Validate("summary", existing, out string err4));
            Assert.IsTrue(err4.Contains("already exists"));

            // Valid name
            Assert.IsTrue(SheetNameValidator.Validate("Quarterly_Report 2026", existing, out string validErr));
            Assert.AreEqual(string.Empty, validErr);
        }

        [TestMethod]
        public void SheetNameValidator_Sanitize_ReplacesIllegalChars()
        {
            string sanitized = SheetNameValidator.Sanitize("Sales/2026*Q1:Final");
            Assert.AreEqual("Sales_2026_Q1_Final", sanitized);

            string history = SheetNameValidator.Sanitize("History");
            Assert.AreEqual("Sheet", history);
        }

        [TestMethod]
        public void ViewModel_LoadsWorkbooksAndWorksheets()
        {
            var ctx = new MockExcelContext();
            var vm = new WorkbookSheetNavigatorViewModel(ctx);

            Assert.AreEqual(2, vm.Workbooks.Count);
            Assert.IsNotNull(vm.SelectedWorkbook);
            Assert.AreEqual("Workbook1.xlsx", vm.SelectedWorkbook.Name);

            Assert.AreEqual(3, vm.Worksheets.Count);
            Assert.IsNotNull(vm.SelectedWorksheet);
            Assert.AreEqual("Sheet1", vm.SelectedWorksheet.Name);
        }

        [TestMethod]
        public void ViewModel_FiltersWorksheetsBySearchText()
        {
            var ctx = new MockExcelContext();
            ctx.Sheets.Clear();
            ctx.Sheets.AddRange(new[] { "Summary_2026", "Summary_2025", "RawData", "Tax_Analysis" });
            var vm = new WorkbookSheetNavigatorViewModel(ctx);

            Assert.AreEqual(4, vm.Worksheets.Count);

            vm.SearchText = "summary";
            Assert.AreEqual(2, vm.Worksheets.Count);
            Assert.IsTrue(vm.Worksheets.All(s => s.Name.StartsWith("Summary", StringComparison.OrdinalIgnoreCase)));

            vm.SearchText = "tax";
            Assert.AreEqual(1, vm.Worksheets.Count);
            Assert.AreEqual("Tax_Analysis", vm.Worksheets[0].Name);

            vm.SearchText = "";
            Assert.AreEqual(4, vm.Worksheets.Count);
        }

        [TestMethod]
        public void ViewModel_FiltersWorksheetsByVisibility()
        {
            var ctx = new MockExcelContext();
            ctx.Sheets.Clear();
            ctx.Sheets.AddRange(new[] { "Sheet1", "Sheet2", "Sheet3" });
            ctx.SheetVisibility["Sheet1"] = true;
            ctx.SheetVisibility["Sheet2"] = false;
            ctx.SheetVisibility["Sheet3"] = true;

            var vm = new WorkbookSheetNavigatorViewModel(ctx);

            vm.FilterMode = SheetFilterMode.VisibleOnly;
            Assert.AreEqual(2, vm.Worksheets.Count);
            Assert.IsTrue(vm.Worksheets.All(s => s.IsVisible));

            vm.FilterMode = SheetFilterMode.HiddenOnly;
            Assert.AreEqual(1, vm.Worksheets.Count);
            Assert.AreEqual("Sheet2", vm.Worksheets[0].Name);

            vm.FilterMode = SheetFilterMode.All;
            Assert.AreEqual(3, vm.Worksheets.Count);
        }

        [TestMethod]
        public void ViewModel_SortsWorkbooks()
        {
            var ctx = new MockExcelContext();
            ctx.MockWorkbookSummaries.Clear();
            ctx.MockWorkbookSummaries.Add(new WorkbookSummaryInfo("Zulu.xlsx", @"C:\Z.xlsx", false, true, 1));
            ctx.MockWorkbookSummaries.Add(new WorkbookSummaryInfo("Alpha.xlsx", @"C:\A.xlsx", true, true, 1));
            ctx.MockWorkbookSummaries.Add(new WorkbookSummaryInfo("Bravo.xlsx", @"C:\B.xlsx", false, true, 1));

            var vm = new WorkbookSheetNavigatorViewModel(ctx);

            vm.WorkbookSort = WorkbookSortOrder.NameAscending;
            Assert.AreEqual("Alpha.xlsx", vm.Workbooks[0].Name);
            Assert.AreEqual("Bravo.xlsx", vm.Workbooks[1].Name);
            Assert.AreEqual("Zulu.xlsx", vm.Workbooks[2].Name);

            vm.WorkbookSort = WorkbookSortOrder.NameDescending;
            Assert.AreEqual("Zulu.xlsx", vm.Workbooks[0].Name);
            Assert.AreEqual("Bravo.xlsx", vm.Workbooks[1].Name);
            Assert.AreEqual("Alpha.xlsx", vm.Workbooks[2].Name);

            vm.WorkbookSort = WorkbookSortOrder.OpenOrder;
            Assert.AreEqual("Zulu.xlsx", vm.Workbooks[0].Name);
            Assert.AreEqual("Alpha.xlsx", vm.Workbooks[1].Name);
        }

        [TestMethod]
        public void ViewModel_MovesWorksheetUpAndDown()
        {
            var ctx = new MockExcelContext();
            ctx.Sheets.Clear();
            ctx.Sheets.AddRange(new[] { "First", "Second", "Third" });
            var vm = new WorkbookSheetNavigatorViewModel(ctx);

            // Select "Second"
            vm.SelectedWorksheet = vm.Worksheets.First(s => s.Name == "Second");

            // Move Up -> becomes index 0
            vm.MoveWorksheet(SheetMoveDirection.Up);
            Assert.AreEqual("Second", ctx.Sheets[0]);
            Assert.AreEqual("First", ctx.Sheets[1]);

            // Move Down -> returns to index 1
            vm.SelectedWorksheet = vm.Worksheets.First(s => s.Name == "Second");
            vm.MoveWorksheet(SheetMoveDirection.Down);
            Assert.AreEqual("First", ctx.Sheets[0]);
            Assert.AreEqual("Second", ctx.Sheets[1]);

            // Move to End
            vm.SelectedWorksheet = vm.Worksheets.First(s => s.Name == "First");
            vm.MoveWorksheet(SheetMoveDirection.End);
            Assert.AreEqual("First", ctx.Sheets.Last());

            // Move to Top
            vm.SelectedWorksheet = vm.Worksheets.First(s => s.Name == "First");
            vm.MoveWorksheet(SheetMoveDirection.Top);
            Assert.AreEqual("First", ctx.Sheets.First());
        }

        [TestMethod]
        public void ViewModel_TogglesLastActiveSheet()
        {
            var ctx = new MockExcelContext();
            ctx.Sheets.Clear();
            ctx.Sheets.AddRange(new[] { "Alpha", "Beta", "Gamma" });
            ctx.ActiveWorksheetName = "Alpha";
            var vm = new WorkbookSheetNavigatorViewModel(ctx);

            // Activate Beta
            var beta = vm.Worksheets.First(s => s.Name == "Beta");
            vm.ActivateWorksheet(beta);
            Assert.AreEqual("Beta", ctx.ActiveWorksheetName);

            // Toggle back to Alpha
            vm.ToggleLastActiveSheet();
            Assert.AreEqual("Alpha", ctx.ActiveWorksheetName);

            // Toggle again to Beta
            vm.ToggleLastActiveSheet();
            Assert.AreEqual("Beta", ctx.ActiveWorksheetName);
        }

        [TestMethod]
        public void ViewModel_AddsAndRenamesAndDeletesWorksheets()
        {
            var ctx = new MockExcelContext();
            ctx.Sheets.Clear();
            ctx.Sheets.AddRange(new[] { "Sheet1", "Sheet2" });
            var vm = new WorkbookSheetNavigatorViewModel(ctx);

            // Add sheet
            vm.AddNewWorksheet("NewMetrics");
            Assert.IsTrue(ctx.Sheets.Contains("NewMetrics"));

            // Rename sheet
            vm.RenameWorksheet("NewMetrics", "Q1_Metrics");
            Assert.IsFalse(ctx.Sheets.Contains("NewMetrics"));
            Assert.IsTrue(ctx.Sheets.Contains("Q1_Metrics"));

            // Delete sheet
            var q1 = vm.Worksheets.First(s => s.Name == "Q1_Metrics");
            vm.DeleteWorksheet(q1);
            Assert.IsFalse(ctx.Sheets.Contains("Q1_Metrics"));
        }

        [TestMethod]
        public void ViewModel_SavesAndClosesWorkbooks()
        {
            var ctx = new MockExcelContext();
            var vm = new WorkbookSheetNavigatorViewModel(ctx);

            var unsaved = vm.Workbooks.First(w => !w.IsSaved);
            vm.SaveWorkbook(unsaved);
            Assert.IsTrue(ctx.MockWorkbookSummaries.First(w => w.Name == unsaved.Name).IsSaved);

            vm.SaveAllWorkbooks();
            Assert.IsTrue(ctx.MockWorkbookSummaries.All(w => w.IsSaved));

            int countBefore = vm.Workbooks.Count;
            vm.CloseWorkbook(vm.Workbooks[0]);
            Assert.AreEqual(countBefore - 1, vm.Workbooks.Count);
        }
    }
}
