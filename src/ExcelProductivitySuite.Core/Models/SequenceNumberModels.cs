using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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
        Horizontal,

        /// <summary>
        /// Fills vertically down first column, then up second column, zig-zagging column by column.
        /// </summary>
        SerpentineVertical,

        /// <summary>
        /// Fills horizontally across first row (left-to-right), then right-to-left across second row, zig-zagging row by row.
        /// </summary>
        SerpentineHorizontal
    }

    /// <summary>
    /// Specifies the numbering format and character set used for sequential generation.
    /// </summary>
    public enum SequenceType
    {
        /// <summary>
        /// Standard numeric progression: 1, 2, 3... or zero-padded 0001, 0002...
        /// </summary>
        Numeric,

        /// <summary>
        /// Uppercase letters: A, B, C... Z, AA, AB...
        /// </summary>
        AlphaUpper,

        /// <summary>
        /// Lowercase letters: a, b, c... z, aa, ab...
        /// </summary>
        AlphaLower,

        /// <summary>
        /// Uppercase Roman numerals: I, II, III, IV, V...
        /// </summary>
        RomanUpper,

        /// <summary>
        /// Lowercase Roman numerals: i, ii, iii, iv, v...
        /// </summary>
        RomanLower,

        /// <summary>
        /// Hexadecimal sequence: 0001, 0002... 000F, 0010...
        /// </summary>
        Hexadecimal,

        /// <summary>
        /// Custom user-defined list of repeating items (e.g. Q1, Q2, Q3, Q4 or Spring, Summer, Autumn, Winter).
        /// </summary>
        CustomList
    }

    /// <summary>
    /// Represents a customizable sequential numbering rule or pattern.
    /// </summary>
    public class SequenceNumberDefinition
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Name { get; set; } = "New Sequence";
        public SequenceType Type { get; set; } = SequenceType.Numeric;
        public List<string> CustomListItems { get; set; } = new List<string>();
        public long StartNumber { get; set; } = 1;
        public long? EndNumber { get; set; } = null; // Optional end limit for cycling/looping
        public long Increment { get; set; } = 1;
        public int Digits { get; set; } = 1;         // Zero padding width (e.g. 4 -> 0001)
        public string Prefix { get; set; } = string.Empty;
        public string Suffix { get; set; } = string.Empty;
        public int RepeatCount { get; set; } = 1;    // Repeat each number N times before incrementing
        public long CurrentNumber { get; set; } = 1; // Tracks current progression state (Next number)
        public bool IsBuiltIn { get; set; } = false;
        public bool UseDynamicPlaceholders { get; set; } = true;
        public bool OutputAsNumeric { get; set; } = false; // When true and sequence is numeric, sets native number in Excel

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
            bool isBuiltIn = false,
            SequenceType type = SequenceType.Numeric,
            List<string>? customListItems = null,
            bool useDynamicPlaceholders = true,
            bool outputAsNumeric = false)
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
            Type = type;
            CustomListItems = customListItems != null ? new List<string>(customListItems) : new List<string>();
            UseDynamicPlaceholders = useDynamicPlaceholders;
            OutputAsNumeric = outputAsNumeric;
        }

        /// <summary>
        /// Resets the current progression counter back to the configured start number.
        /// </summary>
        public void Reset()
        {
            CurrentNumber = StartNumber;
        }

        /// <summary>
        /// Formats a numeric value with prefix, character padding, and suffix using current time.
        /// </summary>
        public string FormatValue(long value)
        {
            return FormatValue(value, null, null);
        }

        /// <summary>
        /// Formats a value with prefix, padding, suffix, and dynamic placeholder tokens (date, sheet).
        /// </summary>
        public string FormatValue(long value, DateTime? referenceDate = null, string? sheetName = null)
        {
            string numStr;
            switch (Type)
            {
                case SequenceType.AlphaUpper:
                    string alphaU = NumberToAlpha(value, lower: false);
                    numStr = Digits > alphaU.Length ? alphaU.PadLeft(Digits, 'A') : alphaU;
                    break;

                case SequenceType.AlphaLower:
                    string alphaL = NumberToAlpha(value, lower: true);
                    numStr = Digits > alphaL.Length ? alphaL.PadLeft(Digits, 'a') : alphaL;
                    break;

                case SequenceType.RomanUpper:
                    numStr = NumberToRoman(value, lower: false);
                    break;

                case SequenceType.RomanLower:
                    numStr = NumberToRoman(value, lower: true);
                    break;

                case SequenceType.Hexadecimal:
                    string hex = Math.Abs(value).ToString("X");
                    if (Digits > hex.Length) hex = hex.PadLeft(Digits, '0');
                    if (value < 0) hex = "-" + hex;
                    numStr = hex;
                    break;

                case SequenceType.CustomList:
                    if (CustomListItems != null && CustomListItems.Count > 0)
                    {
                        int count = CustomListItems.Count;
                        int idx = (int)((value - 1) % count);
                        if (idx < 0) idx += count;
                        numStr = CustomListItems[idx];
                    }
                    else
                    {
                        numStr = value.ToString();
                    }
                    break;

                case SequenceType.Numeric:
                default:
                    numStr = Digits > 1
                        ? Math.Abs(value).ToString().PadLeft(Digits, '0')
                        : Math.Abs(value).ToString();
                    if (value < 0) numStr = "-" + numStr;
                    break;
            }

            string pre = Prefix ?? string.Empty;
            string suf = Suffix ?? string.Empty;

            if (UseDynamicPlaceholders)
            {
                pre = FormatDynamicTokens(pre, referenceDate, sheetName);
                suf = FormatDynamicTokens(suf, referenceDate, sheetName);
            }

            return $"{pre}{numStr}{suf}";
        }

        /// <summary>
        /// Converts a 1-based positive index to alphabetic sequence (1 -> A, 26 -> Z, 27 -> AA).
        /// </summary>
        public static string NumberToAlpha(long value, bool lower = false)
        {
            if (value <= 0) return lower ? "a" : "A";
            var sb = new StringBuilder();
            long rem = value;
            while (rem > 0)
            {
                rem--;
                sb.Insert(0, (char)('A' + (rem % 26)));
                rem /= 26;
            }
            string result = sb.ToString();
            return lower ? result.ToLowerInvariant() : result;
        }

        /// <summary>
        /// Converts an alphabetic string (e.g. "A", "Z", "AA") to its 1-based integer index.
        /// </summary>
        public static long AlphaToNumber(string? alpha)
        {
            if (string.IsNullOrWhiteSpace(alpha)) return 1;
            long result = 0;
            string clean = alpha!.Trim().ToUpperInvariant();
            foreach (char c in clean)
            {
                if (c >= 'A' && c <= 'Z')
                {
                    result = result * 26 + (c - 'A' + 1);
                }
            }
            return result > 0 ? result : 1;
        }

        /// <summary>
        /// Converts a positive integer to Roman numerals (e.g. 1 -> I, 4 -> IV, 9 -> IX, 1999 -> MCMXCIX).
        /// </summary>
        public static string NumberToRoman(long number, bool lower = false)
        {
            if (number <= 0) return "0";
            if (number > 3999) return number.ToString();

            int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] numerals = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };

            var sb = new StringBuilder();
            long rem = number;
            for (int i = 0; i < values.Length; i++)
            {
                while (rem >= values[i])
                {
                    rem -= values[i];
                    sb.Append(numerals[i]);
                }
            }
            string roman = sb.ToString();
            return lower ? roman.ToLowerInvariant() : roman;
        }

        /// <summary>
        /// Converts a Roman numeral string (e.g. "IV", "X", "MCMXCIX") to integer.
        /// </summary>
        public static long RomanToNumber(string? roman)
        {
            if (string.IsNullOrWhiteSpace(roman)) return 1;
            string clean = roman!.Trim().ToUpperInvariant();
            var map = new Dictionary<char, int>
            {
                { 'I', 1 }, { 'V', 5 }, { 'X', 10 }, { 'L', 50 },
                { 'C', 100 }, { 'D', 500 }, { 'M', 1000 }
            };
            long total = 0;
            int prev = 0;
            for (int i = clean.Length - 1; i >= 0; i--)
            {
                if (!map.TryGetValue(clean[i], out int current))
                    continue;
                if (current < prev)
                    total -= current;
                else
                    total += current;
                prev = current;
            }
            return total > 0 ? total : 1;
        }

        /// <summary>
        /// Evaluates dynamic date, time, and context tokens in a string.
        /// Supports {YYYY}, {YY}, {MM}, {M}, {DD}, {D}, {SHEET}, {DATE:format}, {RANDOM:len}.
        /// Also supports bracket syntax [YYYY], [MM], [DD], [DATE:...].
        /// </summary>
        public static string FormatDynamicTokens(string text, DateTime? date = null, string? sheetName = null)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            DateTime d = date ?? DateTime.Now;
            string sheet = sheetName ?? string.Empty;

            string result = text;
            result = Regex.Replace(result, @"\{YYYY\}|\[YYYY\]", d.ToString("yyyy"), RegexOptions.IgnoreCase);
            result = Regex.Replace(result, @"\{YY\}|\[YY\]", d.ToString("yy"), RegexOptions.IgnoreCase);
            result = Regex.Replace(result, @"\{MM\}|\[MM\]", d.ToString("MM"), RegexOptions.IgnoreCase);
            result = Regex.Replace(result, @"\{M\}|\[M\]", d.Month.ToString(), RegexOptions.IgnoreCase);
            result = Regex.Replace(result, @"\{DD\}|\[DD\]", d.ToString("dd"), RegexOptions.IgnoreCase);
            result = Regex.Replace(result, @"\{D\}|\[D\]", d.Day.ToString(), RegexOptions.IgnoreCase);
            result = Regex.Replace(result, @"\{SHEET\}|\[SHEET\]", sheet, RegexOptions.IgnoreCase);

            // Custom date format {DATE:yyyyMMdd} or [DATE:yyyyMMdd]
            result = Regex.Replace(result, @"\{DATE:([^}]+)\}|\[DATE:([^\]]+)\]", m =>
            {
                string fmt = !string.IsNullOrEmpty(m.Groups[1].Value) ? m.Groups[1].Value : m.Groups[2].Value;
                try { return d.ToString(fmt); } catch { return m.Value; }
            }, RegexOptions.IgnoreCase);

            // Random digits {RANDOM:N} or [RANDOM:N]
            result = Regex.Replace(result, @"\{RANDOM:(\d+)\}|\[RANDOM:(\d+)\]", m =>
            {
                string lenStr = !string.IsNullOrEmpty(m.Groups[1].Value) ? m.Groups[1].Value : m.Groups[2].Value;
                if (int.TryParse(lenStr, out int len) && len > 0 && len <= 10)
                {
                    long min = (long)Math.Pow(10, len - 1);
                    long max = (long)Math.Pow(10, len) - 1;
                    return new Random().Next((int)min, (int)max).ToString();
                }
                return m.Value;
            }, RegexOptions.IgnoreCase);

            return result;
        }

        /// <summary>
        /// Generates sample preview items starting from the StartNumber.
        /// </summary>
        public List<string> GeneratePreview(int count = 6, DateTime? date = null, string? sheetName = null)
        {
            var list = new List<string>();
            long val = StartNumber;
            int rep = 0;

            for (int i = 0; i < count; i++)
            {
                list.Add(FormatValue(val, date, sheetName));
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
                Type = Type,
                CustomListItems = new List<string>(CustomListItems ?? Enumerable.Empty<string>()),
                StartNumber = StartNumber,
                EndNumber = EndNumber,
                Increment = Increment,
                Digits = Digits,
                Prefix = Prefix,
                Suffix = Suffix,
                RepeatCount = RepeatCount,
                CurrentNumber = CurrentNumber,
                IsBuiltIn = IsBuiltIn,
                UseDynamicPlaceholders = UseDynamicPlaceholders,
                OutputAsNumeric = OutputAsNumeric
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
        public bool RandomizeOrder { get; set; } = false;  // Shuffles generated sequence numbers across the selection
        public bool RestartOnValueChange { get; set; } = false; // Restarts sequence counter when adjacent column cell value changes
        public int RestartKeyColumnOffset { get; set; } = -1;   // Offset to key column (-1 = left column, 1 = right column)
        public int StepRows { get; set; } = 1;                  // Fill every N-th row (1 = every row, 2 = every 2nd row)
        public DateTime? ReferenceDate { get; set; } = null;
        public string? WorksheetName { get; set; } = null;
    }
}
