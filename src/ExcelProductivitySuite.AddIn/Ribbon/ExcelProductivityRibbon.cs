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
            // 1. First priority: Load from embedded assembly resource (portable across all deployments)
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("ExcelProductivitySuite.AddIn.Ribbon.ExcelProductivityRibbon.xml"))
            {
                if (stream != null)
                {
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }

            // 2. Second priority: Read Ribbon XML from deployment / assembly directory
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string xmlPath = Path.Combine(baseDir, "Ribbon", "ExcelProductivityRibbon.xml");
            if (File.Exists(xmlPath))
            {
                return File.ReadAllText(xmlPath);
            }

            // 3. Fallback: check relative path from assembly location
            string codeBase = Assembly.GetExecutingAssembly().Location;
            if (!string.IsNullOrEmpty(codeBase))
            {
                string asmDir = Path.GetDirectoryName(codeBase) ?? "";
                string asmXmlPath = Path.Combine(asmDir, "Ribbon", "ExcelProductivityRibbon.xml");
                if (File.Exists(asmXmlPath))
                {
                    return File.ReadAllText(asmXmlPath);
                }
            }

            throw new FileNotFoundException($"Cannot locate Ribbon XML as embedded resource or on disk at '{xmlPath}'.");
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
                case "btnCombineText": return "CMD_TEXT_COMBINE";
                case "btnReverseText": return "CMD_TEXT_REVERSE";
                case "btnRemoveChars": return "CMD_TEXT_REMOVE_CHARS";
                case "btnRemoveByPos": return "CMD_TEXT_REMOVE_BY_POS";
                case "btnTrimSpaces": return "CMD_TEXT_TRIM_SPACES";
                case "btnCleanText": return "CMD_TEXT_CLEAN_NONPRINTING";
                case "btnChangeCase": return "CMD_TEXT_CHANGE_CASE";
                case "btnAddText": return "CMD_TEXT_ADD_AFFIX";
                case "btnConvertNumbers": return "CMD_TEXT_CONVERT_NUMBERS";
                case "btnConvertDates": return "CMD_TEXT_CONVERT_DATES";
                case "btnSpellNumber": return "CMD_TEXT_SPELL_NUMBER";
                case "btnTranspose": return "CMD_RANGE_TRANSPOSE";
                case "btnFlipRange": return "CMD_RANGE_FLIP_HORIZONTAL";
                case "btnSwapRanges": return "CMD_RANGE_SWAP";
                case "btnChangeSign": return "CMD_RANGE_CHANGE_SIGN";
                case "btnRoundValues": return "CMD_RANGE_ROUND_VALUES";
                case "btnRandomSort": return "CMD_RANGE_RANDOM_SORT";
                case "btnMergeKeepData": return "CMD_RANGE_MERGE_KEEP_DATA";
                case "btnInsertBlankRows": return "CMD_RANGE_INSERT_BLANK_ROWS";
                case "btnDeleteBlankRows": return "CMD_RANGE_DELETE_BLANK_ROWS";
                case "btnFindDuplicates": return "CMD_FIND_DUPLICATES";
                case "btnAdvancedFindReplace": return "CMD_ADVANCED_FIND_REPLACE";
                case "btnSelectMaxMin": return "CMD_SELECT_MAX_MIN";
                case "btnRowDifferences": return "CMD_SELECT_ROW_DIFFERENCES";
                case "btnCommentInspector": return "CMD_FIND_BY_COMMENT";
                case "btnFillBlanks": return "CMD_RANGE_FILL_BLANKS";
                case "btnCompareRanges": return "CMD_RANGE_COMPARE";
                case "btnSelectSpecial": return "CMD_SELECT_SPECIAL";
                case "btnMergeRanges": return "CMD_MERGE_RANGES";
                case "btnRemoveDuplicates": return "CMD_DATA_REMOVE_DUPLICATES";
                case "btnAdvancedSort": return "CMD_DATA_ADVANCED_SORT";
                case "btnFilterByColor": return "CMD_DATA_FILTER_BY_COLOR";
                case "btnCountByColor": return "CMD_DATA_COUNT_BY_COLOR";
                case "btnMergeTables": return "CMD_DATA_MERGE_TABLES";
                case "btnExtractDistinct": return "CMD_DATA_EXTRACT_DISTINCT";
                case "btnPreventDuplicate": return "CMD_VALIDATION_PREVENT_DUPLICATE";
                case "btnFindMissing": return "CMD_FIND_MISSING_RECORDS";
                case "btnAdvancedFilter": return "CMD_ADVANCED_FILTER";
                case "btnFormulaToValue": return "CMD_FORMULA_TO_VALUE";
                case "btnSheetVisibility": return "CMD_SHEET_VISIBILITY";
                case "btnBatchRename": return "CMD_SHEET_BATCH_RENAME";
                case "btnCombineSheets": return "CMD_SHEET_COMBINE";
                case "btnSortSheetTabs": return "CMD_SHEET_SORT_TABS";
                case "btnCreateSheetIndex": return "CMD_SHEET_CREATE_INDEX";
                case "btnDeleteBlankSheets": return "CMD_SHEET_DELETE_BLANKS";
                case "btnCombineWorkbooksFiles": return "CMD_WB_COMBINE_FILES";
                case "btnSplitWorkbookFiles": return "CMD_WB_SPLIT_FILES";
                case "btnSplitDataToSheets": return "CMD_SPLIT_DATA_TO_SHEETS";
                case "btnCompareWorkbooks": return "CMD_COMPARE_WORKBOOKS";
                case "btnFormulaConvertRef": return "CMD_FORMULA_CONVERT_REF";
                case "btnFormulaExactCopy": return "CMD_FORMULA_EXACT_COPY";
                case "btnFormulaLocateErrors": return "CMD_FORMULA_LOCATE_ERRORS";
                case "btnFormulaAuditPrecedents": return "CMD_FORMULA_AUDIT_PRECEDENTS";
                case "btnFormulaAuditDependents": return "CMD_FORMULA_AUDIT_DEPENDENTS";
                case "btnFormulaReplaceRef": return "CMD_FORMULA_BATCH_REPLACE_REF";
                case "btnFormulaTemplates": return "CMD_FORMULA_TEMPLATES";
                case "btnBatchNumberFormats": return "CMD_FORMAT_BATCH_NUMBERS";
                case "btnDateFormatStandards": return "CMD_FORMAT_DATE_STANDARDS";
                case "btnRangeMath": return "CMD_RANGE_MATH_OPERATIONS";
                case "btnCopyPageSetup": return "CMD_FORMAT_PAGE_SETUP_COPY";
                case "btnFreezePanesMulti": return "CMD_FORMAT_FREEZE_PANES";
                case "btnImportCsv": return "CMD_FILE_IMPORT_CSV";
                case "btnExportRange": return "CMD_FILE_EXPORT_RANGE_CSV";
                case "btnExportPdf": return "CMD_FILE_EXPORT_SHEETS_PDF";
                case "btnFolderListing": return "CMD_FILE_FOLDER_LISTING";
                case "btnInsertTextFile": return "CMD_FILE_INSERT_TEXT_FILE";
                case "btnBatchConvert": return "CMD_FILE_BATCH_CONVERT";
                case "btnQuickStats": return "CMD_ANALYSIS_QUICK_STATS";
                case "btnTargetVsActual": return "CMD_CHART_TARGET_VS_ACTUAL";
                case "btnBatchProtect": return "CMD_SECURITY_BATCH_PROTECT";
                case "btnBatchUnprotect": return "CMD_SECURITY_BATCH_UNPROTECT";
                case "btnCreateSnapshot": return "CMD_RECOVERY_CREATE_SNAPSHOT";
                case "btnRestoreSnapshot": return "CMD_RECOVERY_RESTORE_SNAPSHOT";
                case "btnNamedRanges": return "CMD_NAV_NAMED_RANGES";
                case "btnColumnNav": return "CMD_NAV_COLUMNS";
                case "btnCellPreview": return "CMD_NAV_CELL_PREVIEW";
                case "btnDatePicker": return "CMD_NAV_DATE_PICKER";
                case "btnGridFocus": return "CMD_NAV_GRID_FOCUS";
                case "btnDataEntryForm": return "CMD_NAV_DATA_ENTRY_FORM";
                case "btnAbout": return "CMD_ABOUT";
                default: return controlId;
            }
        }
    }
}
