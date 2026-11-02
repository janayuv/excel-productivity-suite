using System;
using ExcelProductivitySuite.AddIn.ExcelIntegration;
using ExcelProductivitySuite.AddIn.Navigation;
using ExcelProductivitySuite.AddIn.Ribbon;
using ExcelProductivitySuite.AddIn.Startup;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Services;
using Office = Microsoft.Office.Core;

namespace ExcelProductivitySuite.AddIn
{
    public partial class ThisAddIn
    {
        private ICommandRegistry? _registry;
        private IExcelContext? _context;
        private ExcelProductivityRibbon? _ribbon;
        private ExcelKeyboardHook? _keyboardHook;
        private Microsoft.Office.Tools.CustomTaskPane? _navTaskPane;

        public void ToggleNavigationPane()
        {
            try
            {
                if (_navTaskPane == null)
                {
                    if (_context == null) return;
                    var control = new WorkbookSheetNavigatorTaskPaneControl(_context);
                    _navTaskPane = this.CustomTaskPanes.Add(control, "Workbook & Sheet Navigator");
                    _navTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                    _navTaskPane.Width = 340;
                    _navTaskPane.Visible = true;
                }
                else
                {
                    _navTaskPane.Visible = !_navTaskPane.Visible;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError($"[ExcelProductivitySuite] ToggleNavigationPane failed: {ex}");
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            try
            {
                _registry = new CommandRegistry();
                var adapter = new ExcelContextAdapter(() => this.Application);
                adapter.NavigationPaneToggleAction = ToggleNavigationPane;
                _context = adapter;
                _ribbon = new ExcelProductivityRibbon(_registry, _context);
                ToolRegistrar.RegisterAllTools(_registry);
                return _ribbon;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError($"[ExcelProductivitySuite] CreateRibbonExtensibilityObject failed: {ex}");
                throw;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (_registry != null && _registry.GetAllCommands().Count == 0)
            {
                ToolRegistrar.RegisterAllTools(_registry);
            }

            if (_context != null)
            {
                try
                {
                    _keyboardHook = new ExcelKeyboardHook(_context, () => this.Application);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Trace.TraceWarning($"[ExcelProductivitySuite] Keyboard hook setup warning: {ex.Message}");
                }

                try
                {
                    this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
                }
                catch { }
            }
        }

        private void Application_WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook wb, ref bool cancel)
        {
            try
            {
                if (!cancel && wb != null && _context != null)
                {
                    _context.UndoManager.ClearForWorkbook(wb.Name);
                }
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            }
            catch { }

            _keyboardHook?.Dispose();
            _keyboardHook = null;
            _navTaskPane = null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
