using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
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
        private const string IconResourcePrefix = "ExcelProductivitySuite.AddIn.Ribbon.Icons.";
        private const int PicTypeBitmap = 1;
        private const int BiRgb = 0;

        private Office.IRibbonUI? _ribbon;
        private readonly ICommandRegistry _registry;
        private readonly IExcelContext _context;
        private readonly Dictionary<string, stdole.IPictureDisp> _imageCache =
            new Dictionary<string, stdole.IPictureDisp>(StringComparer.OrdinalIgnoreCase);

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

        /// <summary>
        /// Resolves embedded Ribbon PNG icons into alpha-preserving <see cref="stdole.IPictureDisp"/> instances.
        /// </summary>
        public stdole.IPictureDisp? LoadImage(string imageName)
        {
            if (string.IsNullOrWhiteSpace(imageName))
                return null;

            string key = imageName.Trim();
            if (_imageCache.TryGetValue(key, out var cached))
                return cached;

            string resourceName = ResolveIconResourceName(key);
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    return null;

                using (var raw = new Bitmap(stream))
                using (var pargb = raw.PixelFormat == PixelFormat.Format32bppPArgb
                    ? (Bitmap)raw.Clone()
                    : raw.Clone(new Rectangle(0, 0, raw.Width, raw.Height), PixelFormat.Format32bppPArgb))
                {
                    var picture = ConvertToPictureDisp(pargb);
                    if (picture != null)
                        _imageCache[key] = picture;
                    return picture;
                }
            }
        }

        internal static string ResolveIconResourceName(string imageName)
        {
            string normalized = imageName.Trim().Replace('/', '.').Replace('\\', '.').TrimStart('.');
            if (normalized.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                normalized = normalized.Substring(0, normalized.Length - 4);
            if (!normalized.Contains("."))
                normalized = "large." + normalized;
            return $"{IconResourcePrefix}{normalized}.png";
        }

        private static stdole.IPictureDisp? ConvertToPictureDisp(Bitmap bitmap)
        {
            var bmi = new BitmapInfoHeader
            {
                biSize = Marshal.SizeOf(typeof(BitmapInfoHeader)),
                biWidth = bitmap.Width,
                biHeight = -bitmap.Height,
                biPlanes = 1,
                biBitCount = 32,
                biCompression = BiRgb
            };

            IntPtr hBitmap = CreateDIBSection(IntPtr.Zero, ref bmi, 0, out IntPtr bitsPtr, IntPtr.Zero, 0);
            if (hBitmap == IntPtr.Zero || bitsPtr == IntPtr.Zero)
                return null;

            bool ownershipTransferred = false;
            try
            {
                var rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
                var data = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppPArgb);
                try
                {
                    int rowBytes = bitmap.Width * 4;
                    var buffer = new byte[rowBytes];
                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        IntPtr srcRow = IntPtr.Add(data.Scan0, y * data.Stride);
                        IntPtr dstRow = IntPtr.Add(bitsPtr, y * rowBytes);
                        Marshal.Copy(srcRow, buffer, 0, rowBytes);
                        Marshal.Copy(buffer, 0, dstRow, rowBytes);
                    }
                }
                finally
                {
                    bitmap.UnlockBits(data);
                }

                var pictDesc = new PictDescBitmap
                {
                    cbSizeOfStruct = Marshal.SizeOf(typeof(PictDescBitmap)),
                    picType = PicTypeBitmap,
                    hbitmap = hBitmap,
                    hpal = IntPtr.Zero
                };

                Guid iid = typeof(stdole.IPictureDisp).GUID;
                int hr = OleCreatePictureIndirect(ref pictDesc, ref iid, true, out object picture);
                if (hr != 0 || picture == null)
                    return null;

                ownershipTransferred = true;
                return (stdole.IPictureDisp)picture;
            }
            finally
            {
                if (!ownershipTransferred)
                    DeleteObject(hBitmap);
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct BitmapInfoHeader
        {
            public int biSize;
            public int biWidth;
            public int biHeight;
            public short biPlanes;
            public short biBitCount;
            public int biCompression;
            public int biSizeImage;
            public int biXPelsPerMeter;
            public int biYPelsPerMeter;
            public int biClrUsed;
            public int biClrImportant;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PictDescBitmap
        {
            public int cbSizeOfStruct;
            public int picType;
            public IntPtr hbitmap;
            public IntPtr hpal;
        }

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateDIBSection(
            IntPtr hdc,
            [In] ref BitmapInfoHeader pbmi,
            uint iUsage,
            out IntPtr ppvBits,
            IntPtr hSection,
            uint dwOffset);

        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("oleaut32.dll")]
        private static extern int OleCreatePictureIndirect(
            ref PictDescBitmap pictDesc,
            ref Guid refiid,
            [MarshalAs(UnmanagedType.Bool)] bool fOwn,
            [MarshalAs(UnmanagedType.Interface)] out object ppvObj);

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
                case "btnTransformRange": return "CMD_RANGE_TRANSFORM";
                case "btnTranspose": return "CMD_RANGE_TRANSPOSE";
                case "btnFlipRange": return "CMD_RANGE_FLIP_HORIZONTAL";
                case "btnSwapRanges": return "CMD_RANGE_SWAP";
                case "btnChangeSign": return "CMD_RANGE_CHANGE_SIGN";
                case "btnRoundValues": return "CMD_RANGE_ROUND_VALUES";
                case "btnRandomSort": return "CMD_RANGE_RANDOM_SORT";
                case "btnMergeKeepData": return "CMD_RANGE_MERGE_KEEP_DATA";
                case "btnInsertBlankRows": return "CMD_RANGE_INSERT_BLANK_ROWS";
                case "btnDeleteBlankRows": return "CMD_RANGE_DELETE_BLANK_ROWS";
                case "btnDeleteAllEmptyRows": return "CMD_DELETE_ALL_EMPTY_ROWS";
                case "btnBatchInsertCheckboxes": return "CMD_RANGE_INSERT_CHECKBOXES";
                case "btnBatchDeleteCheckboxes": return "CMD_RANGE_DELETE_CHECKBOXES";
                case "btnFindDuplicates": return "CMD_FIND_DUPLICATES";
                case "btnAdvancedFindReplace": return "CMD_ADVANCED_FIND_REPLACE";
                case "btnSelectMaxMin": return "CMD_SELECT_MAX_MIN";
                case "btnRowDifferences": return "CMD_SELECT_ROW_DIFFERENCES";
                case "btnCommentInspector": return "CMD_FIND_BY_COMMENT";
                case "btnFillBlanks": return "CMD_RANGE_FILL_BLANKS";
                case "btnInsertSequenceNumber": return "CMD_RANGE_INSERT_SEQUENCE_NUMBER";
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
                case "btnCopyMultipleRanges": return "CMD_RANGE_COPY_NONCONTIGUOUS";
                case "btnPasteVisibleOnly": return "CMD_RANGE_PASTE_VISIBLE_ONLY";
                case "btnUnmergeAndFill": return "CMD_RANGE_UNMERGE_FILL";
                case "btnRandomSample": return "CMD_RANGE_RANDOM_SAMPLE";
                case "btnFindByFormat": return "CMD_FIND_BY_FORMAT";
                case "btnFindByValidation": return "CMD_FIND_BY_VALIDATION";
                case "btnCopyPageSetup": return "CMD_FORMAT_PAGE_SETUP_COPY";
                case "btnFreezePanesMulti": return "CMD_FORMAT_FREEZE_PANES";
                case "btnImportCsv": return "CMD_FILE_IMPORT_CSV";
                case "btnExportRange": return "CMD_FILE_EXPORT_RANGE_CSV";
                case "btnExportPdf": return "CMD_FILE_EXPORT_SHEETS_PDF";
                case "btnPrintWorkbooks": return "CMD_FILE_PRINT_MULTIPLE_WORKBOOKS";
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
                case "btnNavPane":
                case "btnNavPaneSheets": return "CMD_NAV_PANE";
                case "btnColumnNav": return "CMD_NAV_COLUMNS";
                case "btnCellPreview": return "CMD_NAV_CELL_PREVIEW";
                case "btnDatePicker": return "CMD_NAV_DATE_PICKER";
                case "btnGridFocus":
                case "btnGridFocusToggle":
                case "btnGridFocusSheets":
                case "btnGridFocusToggleSheets":
                case "btnReadingView":
                case "btnReadingViewToggle": return "CMD_NAV_GRID_FOCUS";
                case "btnGridFocusSettings":
                case "btnGridFocusSettingsSheets":
                case "btnReadingViewSettings": return "CMD_NAV_GRID_FOCUS_SETTINGS";
                case "btnGridFocusInspector": return "CMD_NAV_GRID_FOCUS_INSPECTOR";
                case "btnDataEntryForm": return "CMD_NAV_DATA_ENTRY_FORM";
                case "btnExtractNumbers": return "CMD_TEXT_EXTRACT_NUMBERS";
                case "btnExtractEmails": return "CMD_TEXT_EXTRACT_EMAILS";
                case "btnSequenceSheets": return "CMD_SHEET_SEQUENCE_NAMES";
                case "btnWbOpenFolder": return "CMD_WB_OPEN_FOLDER";
                case "btnWbCopyPath": return "CMD_WB_COPY_PATH";
                case "btnWbProperties": return "CMD_WB_PROPERTIES";
                case "btnExtractUrls": return "CMD_TEXT_EXTRACT_URLS";
                case "btnBatchReplace": return "CMD_TEXT_BATCH_REPLACE";
                case "btnFrequencyStats": return "CMD_DATA_FREQUENCY_STATS";
                case "btnClearValidation": return "CMD_VALIDATION_BULK_REMOVE";
                case "btnAdjustRowHeight": return "CMD_FORMAT_ROW_HEIGHT";
                case "btnAdjustColWidth": return "CMD_FORMAT_COL_WIDTH";
                case "btnCreateDropdown": return "CMD_VALIDATION_CREATE_DROPDOWN";
                case "btnColumnDifferences": return "CMD_SELECT_COLUMN_DIFFERENCES";
                case "btnSuiteUndo": return "CMD_SUITE_UNDO";
                case "btnSuiteRedo": return "CMD_SUITE_REDO";
                case "btnAbout": return "CMD_ABOUT";
                case "btnCheckUpdates": return "CMD_CHECK_UPDATES";
                default: return controlId;
            }
        }
    }
}
