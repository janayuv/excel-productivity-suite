using System;
using System.Collections.Generic;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Specifies the cell traversal direction when inserting sequence numbers across a range.
    /// </summary>
    public enum SequenceFillOrder
    {
        /// <summary>
        /// Fills vertically down columns (top-to-bottom, then next column).
        /// </summary>
        Vertical,

        /// <summary>
        /// Fills horizontally across rows (left-to-right, then next row).
        /// </summary>
        Horizontal
    }

    /// <summary>
    /// Represents a customizable sequential numbering rule or pattern.
    /// </summary>
    public class SequenceNumberDefinition
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Name { get; set; } = "New Sequence";
        public long StartNumber { get; set; } = 1;
        public long? EndNumber { get; set; } = null; // Optional end limit for cycling/looping
        public long Increment { get; set; } = 1;
        public int Digits { get; set; } = 1;         // Zero padding width (e.g. 4 -> 0001)
        public string Prefix { get; set; } = string.Empty;
        public string Suffix { get; set; } = string.Empty;
        public int RepeatCount { get; set; } = 1;    // Repeat each number N times before incrementing
        public long CurrentNumber { get; set; } = 1; // Tracks current progression state (Next number)
        public bool IsBuiltIn { get; set; } = false;

        public SequenceNumberDefinition()
        {
            CurrentNumber = StartNumber;
        }

        public SequenceNumberDefinition(
            string name,
            long startNumber = 1,
            long increment = 1,
            int digits = 1,
            string prefix = "",
            string suffix = "",
            long? endNumber = null,
            int repeatCount = 1,
            bool isBuiltIn = false)
        {
            Name = name;
            StartNumber = startNumber;
            Increment = increment;
            Digits = Math.Max(1, Math.Min(digits, 12));
            Prefix = prefix ?? string.Empty;
            Suffix = suffix ?? string.Empty;
            EndNumber = endNumber;
            RepeatCount = Math.Max(1, repeatCount);
            CurrentNumber = startNumber;
            IsBuiltIn = isBuiltIn;
        }

        /// <summary>
        /// Resets the current progression counter back to the configured start number.
        /// </summary>
        public void Reset()
        {
            CurrentNumber = StartNumber;
        }

        /// <summary>
        /// Formats a numeric value with prefix, zero padding, and suffix.
        /// </summary>
        public string FormatValue(long value)
        {
            string numStr = Digits > 1
                ? Math.Abs(value).ToString().PadLeft(Digits, '0')
                : Math.Abs(value).ToString();

            if (value < 0) numStr = "-" + numStr;

            return $"{Prefix}{numStr}{Suffix}";
        }

        /// <summary>
        /// Generates sample preview items starting from the StartNumber.
        /// </summary>
        public List<string> GeneratePreview(int count = 6)
        {
            var list = new List<string>();
            long val = StartNumber;
            int rep = 0;

            for (int i = 0; i < count; i++)
            {
                list.Add(FormatValue(val));
                rep++;
                if (rep >= RepeatCount)
                {
                    rep = 0;
                    val += Increment;
                    if (EndNumber.HasValue)
                    {
                        if (Increment > 0 && val > EndNumber.Value) val = StartNumber;
                        else if (Increment < 0 && val < EndNumber.Value) val = StartNumber;
                    }
                }
            }
            return list;
        }

        /// <summary>
        /// Creates a deep copy of this definition.
        /// </summary>
        public SequenceNumberDefinition Clone()
        {
            return new SequenceNumberDefinition
            {
                Id = Id,
                Name = Name,
                StartNumber = StartNumber,
                EndNumber = EndNumber,
                Increment = Increment,
                Digits = Digits,
                Prefix = Prefix,
                Suffix = Suffix,
                RepeatCount = RepeatCount,
                CurrentNumber = CurrentNumber,
                IsBuiltIn = IsBuiltIn
            };
        }
    }

    /// <summary>
    /// Configuration options when inserting sequence numbers into selected cells.
    /// </summary>
    public class SequenceFillOptions
    {
        public SequenceFillOrder FillOrder { get; set; } = SequenceFillOrder.Vertical;
        public bool OnlyFillBlankCells { get; set; } = false;
        public bool SkipHiddenCells { get; set; } = true;
        public bool ContinueFromLast { get; set; } = true; // If true, uses CurrentNumber; if false, starts from StartNumber
    }
}
