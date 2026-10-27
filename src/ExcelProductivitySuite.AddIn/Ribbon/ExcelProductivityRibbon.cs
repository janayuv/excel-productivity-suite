using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelProductivitySuite.Core.Abstractions;
using Office = Microsoft.Office.Core;

namespace ExcelProductivitySuite.AddIn.Ribbon
{
    [ComVisible(true)]
    public class ExcelProductivityRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI? _ribbon;
        private readonly ICommandRegistry _registry;
        private readonly IExcelContext _context;

        public ExcelProductivityRibbon(ICommandRegistry registry, IExcelContext context)
        {
            _registry = registry ?? throw new ArgumentNullException(nameof(registry));
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public string GetCustomUI(string ribbonID)
        {
            // Read Ribbon XML from assembly directory or embedded resource
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string xmlPath = Path.Combine(baseDir, "Ribbon", "ExcelProductivityRibbon.xml");

            if (File.Exists(xmlPath))
            {
                return File.ReadAllText(xmlPath);
            }

            // Fallback: check project relative path during dev
            string devPath = @"d:\Kutools-New\src\ExcelProductivitySuite.AddIn\Ribbon\ExcelProductivityRibbon.xml";
            if (File.Exists(devPath))
            {
                return File.ReadAllText(devPath);
            }

            throw new FileNotFoundException($"Cannot locate Ribbon XML at {xmlPath} or {devPath}");
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void OnRibbonButtonClicked(Office.IRibbonControl control)
        {
            if (control == null) return;

            string commandId = MapControlIdToCommandId(control.Id);
            try
            {
                _registry.Execute(commandId, _context);
            }
            catch (Exception ex)
            {
                _context.ShowMessage($"Error executing command '{control.Id}': {ex.Message}", "Command Error", isError: true);
            }
        }

        private static string MapControlIdToCommandId(string controlId)
        {
            switch (controlId)
            {
                case "btnLauncher": return "CMD_LAUNCHER";
                case "btnExtractText": return "CMD_TEXT_EXTRACT";
                case "btnSplitText": return "CMD_TEXT_SPLIT";
                case "btnRemoveChars": return "CMD_TEXT_REMOVE_CHARS";
                case "btnTrimSpaces": return "CMD_TEXT_TRIM_SPACES";
                case "btnChangeCase": return "CMD_TEXT_CHANGE_CASE";
                case "btnAddText": return "CMD_TEXT_ADD_AFFIX";
                case "btnFillBlanks": return "CMD_RANGE_FILL_BLANKS";
                case "btnCompareRanges": return "CMD_RANGE_COMPARE";
                case "btnSelectSpecial": return "CMD_SELECT_SPECIAL";
                case "btnRemoveDuplicates": return "CMD_DATA_REMOVE_DUPLICATES";
                case "btnFormulaToValue": return "CMD_FORMULA_TO_VALUE";
                case "btnSheetVisibility": return "CMD_SHEET_VISIBILITY";
                case "btnBatchRename": return "CMD_SHEET_BATCH_RENAME";
                case "btnAbout": return "CMD_ABOUT";
                default: return controlId;
            }
        }
    }
}
