using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Specifies the target data scope for PDF export operations.
    /// </summary>
    public enum PdfExportScope
    {
        ActiveSheet,
        Selection,
        EntireWorkbook,
        SelectedSheets
    }

    /// <summary>
    /// Specifies the output structure when exporting one or more worksheets.
    /// </summary>
    public enum PdfOutputMode
    {
        SingleCombinedFile,
        SeparateFilesPerSheet
    }

    /// <summary>
    /// Page orientation overrides for PDF publishing.
    /// </summary>
    public enum PdfPageOrientation
    {
        Default,
        Portrait,
        Landscape
    }

    /// <summary>
    /// Page scaling and fit options for PDF publishing.
    /// </summary>
    public enum PdfPageFit
    {
        Default,
        FitSheetOnOnePage,
        FitAllColumnsOnOnePage,
        FitAllRowsOnOnePage
    }

    /// <summary>
    /// Standard paper sizes for PDF publishing.
    /// </summary>
    public enum PdfPaperSize
    {
        Default,
        A4,
        Letter,
        Legal,
        A3,
        Tabloid,
        Executive
    }

    /// <summary>
    /// PDF output quality.
    /// </summary>
    public enum PdfQuality
    {
        Standard,
        MinimumSize
    }

    /// <summary>
    /// Configuration settings for PDF export operations.
    /// </summary>
    public class PdfExportSettings
    {
        public PdfExportScope Scope { get; set; } = PdfExportScope.ActiveSheet;
        public PdfOutputMode OutputMode { get; set; } = PdfOutputMode.SingleCombinedFile;
        public List<string> SelectedSheets { get; set; } = new();

        /// <summary>
        /// Destination file path for SingleCombinedFile mode.
        /// </summary>
        public string DestinationPath { get; set; } = string.Empty;

        /// <summary>
        /// Output directory for SeparateFilesPerSheet mode.
        /// </summary>
        public string OutputDirectory { get; set; } = string.Empty;

        /// <summary>
        /// Naming pattern for SeparateFilesPerSheet mode.
        /// Tokens: [SheetName], [WorkbookName], [Index], [Date], [DateTime]
        /// </summary>
        public string FileNamePattern { get; set; } = "[WorkbookName] - [SheetName]";

        // Page Range
        public bool ExportPageRange { get; set; } = false;
        public int PageFrom { get; set; } = 1;
        public int PageTo { get; set; } = 1;

        // Quality & Document Options
        public PdfQuality Quality { get; set; } = PdfQuality.Standard;
        public bool IncludeDocumentProperties { get; set; } = true;
        public bool IgnorePrintAreas { get; set; } = false;

        // Page Setup & Layout Overrides
        public PdfPageOrientation Orientation { get; set; } = PdfPageOrientation.Default;
        public PdfPageFit PageFit { get; set; } = PdfPageFit.Default;
        public PdfPaperSize PaperSize { get; set; } = PdfPaperSize.Default;

        // Post-export Actions
        public bool OpenAfterExport { get; set; } = true;
        public bool OpenOutputFolder { get; set; } = false;
    }

    /// <summary>
    /// Represents the outcome of a PDF export operation.
    /// </summary>
    public class PdfExportResult
    {
        public bool Success { get; set; }
        public int FilesExported { get; set; }
        public List<string> ExportedFiles { get; set; } = new();
        public string Message { get; set; } = string.Empty;
        public string? ErrorMessage { get; set; }

        public static PdfExportResult Succeeded(int count, IEnumerable<string> files, string message) =>
            new PdfExportResult
            {
                Success = true,
                FilesExported = count,
                ExportedFiles = new List<string>(files),
                Message = message
            };

        public static PdfExportResult Failed(string errorMessage) =>
            new PdfExportResult
            {
                Success = false,
                FilesExported = 0,
                ErrorMessage = errorMessage,
                Message = errorMessage
            };
    }
}
