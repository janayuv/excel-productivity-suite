using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Defines cell linking modes for batch inserted checkboxes.
    /// Corresponds to Kutools for Excel: Batch Insert Smart Checkboxes.
    /// </summary>
    public enum CheckboxLinkMode
    {
        /// <summary>
        /// Insert visual checkboxes only without linking them to any cell.
        /// </summary>
        None,

        /// <summary>
        /// Each checkbox is linked directly to the cell it sits in.
        /// When toggled, the cell value switches between TRUE and FALSE.
        /// </summary>
        OwnCell,

        /// <summary>
        /// Checkboxes are linked to an external cell range (e.g. adjacent column or specified range).
        /// Keeps the original data clean while maintaining logical control.
        /// </summary>
        OtherCells
    }

    /// <summary>
    /// Configuration options for batch inserting smart checkboxes into a worksheet.
    /// </summary>
    public class BatchInsertCheckboxOptions
    {
        /// <summary>
        /// Target worksheet range address where checkboxes will be inserted (e.g. "$A$2:$A$10").
        /// </summary>
        public string SourceRangeAddress { get; set; } = string.Empty;

        /// <summary>
        /// Cell linking mode (None, OwnCell, or OtherCells).
        /// Default is OwnCell.
        /// </summary>
        public CheckboxLinkMode LinkMode { get; set; } = CheckboxLinkMode.OwnCell;

        /// <summary>
        /// Target destination range address when LinkMode is OtherCells (e.g. "$B$2:$B$10" or "$B$2").
        /// </summary>
        public string? TargetLinkRangeAddress { get; set; }

        /// <summary>
        /// Automatically uses existing cell text as the checkbox caption/label.
        /// Default is false.
        /// </summary>
        public bool UseCellContentAsLabel { get; set; } = false;

        /// <summary>
        /// When enabled, sets the logical value (TRUE/FALSE) in the linked cell so formulas like COUNTIF evaluate immediately.
        /// Default is true.
        /// </summary>
        public bool DisplayLogicalValue { get; set; } = true;

        /// <summary>
        /// Initial state of the inserted checkboxes. Default is false (Unchecked).
        /// </summary>
        public bool InitialChecked { get; set; } = false;

        /// <summary>
        /// When cell content is used as a checkbox label, whether to clear the underlying cell's text
        /// so it does not render underneath the checkbox caption. Default is true.
        /// </summary>
        public bool ClearOriginalContentWhenLabeled { get; set; } = true;
    }

    /// <summary>
    /// Result of a batch checkbox insertion operation.
    /// </summary>
    public class BatchInsertCheckboxResult
    {
        public bool Success { get; }
        public int CheckboxesInserted { get; }
        public string Message { get; }
        public IReadOnlyList<string> AffectedCellAddresses { get; }

        public BatchInsertCheckboxResult(bool success, int count, string message, IReadOnlyList<string>? affectedAddresses = null)
        {
            Success = success;
            CheckboxesInserted = count;
            Message = message ?? string.Empty;
            AffectedCellAddresses = affectedAddresses ?? Array.Empty<string>();
        }

        public static BatchInsertCheckboxResult Successful(int count, string message, IReadOnlyList<string>? affectedAddresses = null)
        {
            return new BatchInsertCheckboxResult(true, count, message, affectedAddresses);
        }

        public static BatchInsertCheckboxResult Failed(string message)
        {
            return new BatchInsertCheckboxResult(false, 0, message, Array.Empty<string>());
        }
    }

    /// <summary>
    /// Scope for batch deleting checkboxes from a worksheet.
    /// </summary>
    public enum CheckboxDeleteScope
    {
        /// <summary>
        /// Delete checkboxes intersecting the currently selected range.
        /// </summary>
        Selection,

        /// <summary>
        /// Delete all checkboxes across the entire active worksheet.
        /// </summary>
        ActiveWorksheet
    }

    /// <summary>
    /// Configuration options for batch deleting checkboxes.
    /// </summary>
    public class BatchDeleteCheckboxOptions
    {
        public CheckboxDeleteScope Scope { get; set; } = CheckboxDeleteScope.Selection;
        public string? RangeAddress { get; set; }
    }

    /// <summary>
    /// Result of a batch checkbox deletion operation.
    /// </summary>
    public class BatchDeleteCheckboxResult
    {
        public bool Success { get; }
        public int CheckboxesDeleted { get; }
        public string Message { get; }

        public BatchDeleteCheckboxResult(bool success, int count, string message)
        {
            Success = success;
            CheckboxesDeleted = count;
            Message = message ?? string.Empty;
        }

        public static BatchDeleteCheckboxResult Successful(int count, string message)
        {
            return new BatchDeleteCheckboxResult(true, count, message);
        }

        public static BatchDeleteCheckboxResult Failed(string message)
        {
            return new BatchDeleteCheckboxResult(false, 0, message);
        }
    }
}
