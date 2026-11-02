using System;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Specifies the primary restructuring direction for the Transform Range tool.
    /// </summary>
    public enum TransformType
    {
        /// <summary>
        /// Converts a single vertical column into a multi-column and multi-row table.
        /// </summary>
        SingleColumnToRange = 0,

        /// <summary>
        /// Converts a single horizontal row into a multi-column and multi-row table.
        /// </summary>
        SingleRowToRange = 1,

        /// <summary>
        /// Flattens a 2D range into a single vertical column.
        /// </summary>
        RangeToSingleColumn = 2,

        /// <summary>
        /// Flattens a 2D range into a single horizontal row.
        /// </summary>
        RangeToSingleRow = 3
    }

    /// <summary>
    /// Specifies how records are partitioned when transforming a list to a range.
    /// </summary>
    public enum TransformRecordSplit
    {
        /// <summary>
        /// Wraps items into records with a fixed number of columns.
        /// </summary>
        FixedValue = 0,

        /// <summary>
        /// Uses empty or blank cells as record boundaries.
        /// </summary>
        BlankCellDelimited = 1
    }

    /// <summary>
    /// Traversal order when flattening a 2D range into a single column or row.
    /// </summary>
    public enum TransformOrder
    {
        /// <summary>
        /// Read row-by-row: across columns left-to-right, then down to next row.
        /// </summary>
        ByRow = 0,

        /// <summary>
        /// Read column-by-column: down rows top-to-bottom, then right to next column.
        /// </summary>
        ByColumn = 1
    }

    /// <summary>
    /// Specifies where the transformed matrix will be placed.
    /// </summary>
    public enum TransformDestinationMode
    {
        /// <summary>
        /// Overwrites starting at the active selection's top-left cell.
        /// </summary>
        InPlace = 0,

        /// <summary>
        /// Pastes starting at a specified cell address on the active worksheet.
        /// </summary>
        TargetCell = 1,

        /// <summary>
        /// Creates a new worksheet and populates the transformed data.
        /// </summary>
        NewSheet = 2
    }

    /// <summary>
    /// Encapsulates all configurable settings for the Transform Range tool.
    /// </summary>
    public sealed class TransformRangeOptions
    {
        /// <summary>
        /// Transformation mode (Column to Range, Row to Range, Range to Column, Range to Row).
        /// </summary>
        public TransformType Type { get; set; } = TransformType.SingleColumnToRange;

        /// <summary>
        /// Record splitting strategy when converting list to range.
        /// </summary>
        public TransformRecordSplit RecordSplit { get; set; } = TransformRecordSplit.FixedValue;

        /// <summary>
        /// Number of columns per record when using FixedValue split. Default is 5.
        /// </summary>
        public int ColumnsPerRecord { get; set; } = 5;

        /// <summary>
        /// Traversal order when flattening a range.
        /// </summary>
        public TransformOrder Order { get; set; } = TransformOrder.ByRow;

        /// <summary>
        /// If true, empty, null, or whitespace-only cells are omitted during range flattening.
        /// </summary>
        public bool SkipBlankCells { get; set; } = false;

        /// <summary>
        /// Destination location mode.
        /// </summary>
        public TransformDestinationMode DestinationMode { get; set; } = TransformDestinationMode.TargetCell;

        /// <summary>
        /// Destination top-left cell address (e.g. "$E$1" or "G2").
        /// </summary>
        public string TargetAddress { get; set; } = "$E$1";

        /// <summary>
        /// Name of the new worksheet when DestinationMode is NewSheet.
        /// </summary>
        public string NewSheetName { get; set; } = "Transformed";

        /// <summary>
        /// Validates settings before executing transformation.
        /// </summary>
        public (bool IsValid, string? ErrorMessage) Validate()
        {
            if (Type == TransformType.SingleColumnToRange || Type == TransformType.SingleRowToRange)
            {
                if (RecordSplit == TransformRecordSplit.FixedValue && ColumnsPerRecord < 1)
                {
                    return (false, "Columns per record must be greater than or equal to 1.");
                }
            }

            if (DestinationMode == TransformDestinationMode.TargetCell)
            {
                if (string.IsNullOrWhiteSpace(TargetAddress))
                {
                    return (false, "Please specify a valid destination cell address (e.g. $E$1).");
                }
            }
            else if (DestinationMode == TransformDestinationMode.NewSheet)
            {
                if (string.IsNullOrWhiteSpace(NewSheetName))
                {
                    return (false, "Please specify a name for the new worksheet.");
                }
            }

            return (true, null);
        }
    }
}
