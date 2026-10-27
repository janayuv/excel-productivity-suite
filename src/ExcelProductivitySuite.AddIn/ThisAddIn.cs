using System;
using ExcelProductivitySuite.AddIn.ExcelIntegration;
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

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _registry = new CommandRegistry();
            _context = new ExcelContextAdapter(this.Application);
            _ribbon = new ExcelProductivityRibbon(_registry, _context);
            return _ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (_registry != null)
            {
                ToolRegistrar.RegisterAllTools(_registry);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clean unhooking and disposal
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
