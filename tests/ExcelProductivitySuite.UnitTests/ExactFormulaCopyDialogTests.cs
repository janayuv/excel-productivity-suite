using System;
using System.Threading;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Formulas;
using ExcelProductivitySuite.UI.Dialogs;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class ExactFormulaCopyDialogTests
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
        public void ExactCopyFormulaDialog_InitializesWithActiveSelectionAndSelectButtons()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                mock.SelectionAddress = "$A$1:$C$10";

                var dlg = new ExactCopyFormulaDialog(mock);

                Assert.AreEqual("$A$1:$C$10", dlg.SourceTextBox.Text);
                Assert.IsNotNull(dlg.SourceSelectButton, "SourceSelectButton should be present in XAML");
                Assert.IsNotNull(dlg.DestinationSelectButton, "DestinationSelectButton should be present in XAML");
                Assert.IsNotNull(dlg.CopyFormattingCheckBox, "CopyFormattingCheckBox should be present in XAML");
                Assert.IsTrue(dlg.CopyFormattingCheckBox.IsChecked == true, "Copy formatting should be checked by default");
            });
        }

        [TestMethod]
        public void ExactCopyFormulaDialog_PromptRangeForTextBox_UpdatesSourceRange()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                mock.SelectionAddress = "$A$1";
                mock.PromptForRangeResult = "$B$2:$D$20";

                var dlg = new ExactCopyFormulaDialog(mock);
                dlg.PromptRangeForTextBox(dlg.SourceTextBox, "Ranges to be copied :");

                Assert.AreEqual("$B$2:$D$20", dlg.SourceTextBox.Text);
                Assert.AreEqual(1, mock.RangePromptHistory.Count);
                Assert.AreEqual("Ranges to be copied :", mock.RangePromptHistory[0].Prompt);
            });
        }

        [TestMethod]
        public void ExactCopyFormulaDialog_PromptRangeForTextBox_UpdatesDestinationRange()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                mock.PromptForRangeResult = "Sheet2!$F$5";

                var dlg = new ExactCopyFormulaDialog(mock);
                dlg.PromptRangeForTextBox(dlg.DestinationTextBox, "Please select a cell or range to paste :");

                Assert.AreEqual("Sheet2!$F$5", dlg.DestinationTextBox.Text);
                Assert.AreEqual(1, mock.RangePromptHistory.Count);
                Assert.AreEqual("Please select a cell or range to paste :", mock.RangePromptHistory[0].Prompt);
            });
        }

        [TestMethod]
        public void ExactCopyFormulaDialog_PromptRangeForTextBox_PreservesValueOnCancel()
        {
            RunOnSta(() =>
            {
                var mock = new MockExcelContext();
                mock.PromptForRangeResult = null; // User cancelled prompt

                var dlg = new ExactCopyFormulaDialog(mock);
                dlg.SourceTextBox.Text = "$K$1:$M$5";

                dlg.PromptRangeForTextBox(dlg.SourceTextBox, "Ranges to be copied :");

                Assert.AreEqual("$K$1:$M$5", dlg.SourceTextBox.Text, "Address should remain intact when prompt is cancelled");
            });
        }

        [TestMethod]
        public void ExactCopyFormula_CopyFormattingHistory_RecordsFormattedSourceAndDestination()
        {
            var mock = new MockExcelContext();
            mock.CurrentFormulas = new object[,]
            {
                { "=SUM(A1:A5)", "=AVERAGE(B1:B5)" }
            };

            // Test mock method directly
            mock.CopyRangeFormatsByAddress("$A$1:$B$5", "$H$1:$I$5");
            Assert.AreEqual(1, mock.CopiedFormatsHistory.Count);
            Assert.AreEqual("$A$1:$B$5", mock.CopiedFormatsHistory[0].Source);
            Assert.AreEqual("$H$1:$I$5", mock.CopiedFormatsHistory[0].Destination);
        }
    }
}
