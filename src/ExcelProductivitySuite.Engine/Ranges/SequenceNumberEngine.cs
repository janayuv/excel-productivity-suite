using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Ranges
{
    /// <summary>
    /// Core engine for generating, formatting, persisting, and calculating sequential numbers across matrices and ranges.
    /// </summary>
    public static class SequenceNumberEngine
    {
        private static readonly string SettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelProductivitySuite");

        private static readonly string SettingsFilePath = Path.Combine(SettingsDir, "SequenceDefinitions.json");

        /// <summary>
        /// Gets the standard built-in sequence presets.
        /// </summary>
        public static List<SequenceNumberDefinition> GetDefaultPresets()
        {
            int currentYear = DateTime.Now.Year;

            return new List<SequenceNumberDefinition>
            {
                new SequenceNumberDefinition("Simple Numbers (1, 2, 3...)", 1, 1, 1, "", "", null, 1, true),
                new SequenceNumberDefinition("Leading Zeros (0001, 0002...)", 1, 1, 4, "", "", null, 1, true),
                new SequenceNumberDefinition("Invoice Code (INV-0001)", 1, 1, 4, "INV-", "", null, 1, true),
                new SequenceNumberDefinition("Order Number (ORD-" + currentYear + "-0001)", 1, 1, 4, $"ORD-{currentYear}-", "", null, 1, true),
                new SequenceNumberDefinition("Employee ID (EMP-001)", 1, 1, 3, "EMP-", "", null, 1, true),
                new SequenceNumberDefinition("Repeating Cycle (1 to 5)", 1, 1, 1, "", "", 5, 1, true),
                new SequenceNumberDefinition("Repeating Pairs (1, 1, 2, 2...)", 1, 1, 1, "", "", null, 2, true),
                new SequenceNumberDefinition("Item Suffix (Item-01-A)", 1, 1, 2, "Item-", "-A", null, 1, true)
            };
        }

        /// <summary>
        /// Generates a list of formatted sequential strings of the specified length.
        /// </summary>
        public static List<string> GenerateSequence(SequenceNumberDefinition def, int count, out long nextNumber)
        {
            return GenerateSequence(def, count, false, out nextNumber);
        }

        /// <summary>
        /// Generates a list of formatted sequential strings of the specified length.
        /// </summary>
        public static List<string> GenerateSequence(SequenceNumberDefinition def, int count, bool continueFromCurrent, out long nextNumber)
        {
            if (def == null) throw new ArgumentNullException(nameof(def));
            count = Math.Max(0, count);

            var result = new List<string>(count);
            long val = continueFromCurrent ? def.CurrentNumber : def.StartNumber;
            int rep = 0;

            for (int i = 0; i < count; i++)
            {
                result.Add(def.FormatValue(val));
                rep++;
                if (rep >= def.RepeatCount)
                {
                    rep = 0;
                    val += def.Increment;
                    if (def.EndNumber.HasValue)
                    {
                        if (def.Increment > 0 && val > def.EndNumber.Value) val = def.StartNumber;
                        else if (def.Increment < 0 && val < def.EndNumber.Value) val = def.StartNumber;
                    }
                }
            }

            nextNumber = val;
            return result;
        }

        /// <summary>
        /// Populates a 2D matrix with sequence numbers based on traversal order and fill criteria.
        /// </summary>
        public static object?[,] FormatMatrix(
            object?[,] matrix,
            SequenceNumberDefinition def,
            SequenceFillOptions options,
            out int updatedCount,
            out int skippedCount,
            out long nextNumber)
        {
            if (matrix == null) throw new ArgumentNullException(nameof(matrix));
            if (def == null) throw new ArgumentNullException(nameof(def));
            options ??= new SequenceFillOptions();

            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            int rows = rEnd - rStart + 1;
            int cols = cEnd - cStart + 1;

            object?[,] output = (object?[,])matrix.Clone();
            int updated = 0;
            int skipped = 0;

            long val = options.ContinueFromLast ? def.CurrentNumber : def.StartNumber;
            int rep = 0;

            void ProcessCell(int r, int c)
            {
                object? existing = output[r, c];
                bool isBlank = existing == null || string.IsNullOrWhiteSpace(existing.ToString());

                if (options.OnlyFillBlankCells && !isBlank)
                {
                    skipped++;
                    return;
                }

                output[r, c] = def.FormatValue(val);
                updated++;

                rep++;
                if (rep >= def.RepeatCount)
                {
                    rep = 0;
                    val += def.Increment;
                    if (def.EndNumber.HasValue)
                    {
                        if (def.Increment > 0 && val > def.EndNumber.Value) val = def.StartNumber;
                        else if (def.Increment < 0 && val < def.EndNumber.Value) val = def.StartNumber;
                    }
                }
            }

            if (options.FillOrder == SequenceFillOrder.Vertical)
            {
                // Down columns first, then next column
                for (int c = cStart; c <= cEnd; c++)
                {
                    for (int r = rStart; r <= rEnd; r++)
                    {
                        ProcessCell(r, c);
                    }
                }
            }
            else
            {
                // Across rows first, then next row
                for (int r = rStart; r <= rEnd; r++)
                {
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        ProcessCell(r, c);
                    }
                }
            }

            updatedCount = updated;
            skippedCount = skipped;
            nextNumber = val;
            return output;
        }

        /// <summary>
        /// Loads saved sequence definitions from user storage or returns default presets if not found.
        /// </summary>
        public static List<SequenceNumberDefinition> LoadSequences()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    string json = File.ReadAllText(SettingsFilePath, Encoding.UTF8);
                    var list = DeserializeJson(json);
                    if (list != null && list.Count > 0)
                    {
                        return list;
                    }
                }
            }
            catch { }

            return GetDefaultPresets();
        }

        /// <summary>
        /// Saves custom sequence definitions to persistent user storage.
        /// </summary>
        public static void SaveSequences(IEnumerable<SequenceNumberDefinition> sequences)
        {
            if (sequences == null) return;

            try
            {
                if (!Directory.Exists(SettingsDir))
                {
                    Directory.CreateDirectory(SettingsDir);
                }

                string json = SerializeJson(sequences.ToList());
                File.WriteAllText(SettingsFilePath, json, Encoding.UTF8);
            }
            catch { }
        }

        private static string SerializeJson(List<SequenceNumberDefinition> list)
        {
            var sb = new StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < list.Count; i++)
            {
                var s = list[i];
                sb.AppendLine("  {");
                sb.AppendLine($"    \"Id\": \"{Escape(s.Id)}\",");
                sb.AppendLine($"    \"Name\": \"{Escape(s.Name)}\",");
                sb.AppendLine($"    \"StartNumber\": {s.StartNumber},");
                sb.AppendLine($"    \"EndNumber\": {(s.EndNumber.HasValue ? s.EndNumber.Value.ToString() : "null")},");
                sb.AppendLine($"    \"Increment\": {s.Increment},");
                sb.AppendLine($"    \"Digits\": {s.Digits},");
                sb.AppendLine($"    \"Prefix\": \"{Escape(s.Prefix)}\",");
                sb.AppendLine($"    \"Suffix\": \"{Escape(s.Suffix)}\",");
                sb.AppendLine($"    \"RepeatCount\": {s.RepeatCount},");
                sb.AppendLine($"    \"CurrentNumber\": {s.CurrentNumber},");
                sb.AppendLine($"    \"IsBuiltIn\": {(s.IsBuiltIn ? "true" : "false")}");
                sb.Append("  }");
                if (i < list.Count - 1) sb.Append(",");
                sb.AppendLine();
            }
            sb.AppendLine("]");
            return sb.ToString();
        }

        private static List<SequenceNumberDefinition>? DeserializeJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;

            var list = new List<SequenceNumberDefinition>();
            var lines = json.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            SequenceNumberDefinition? current = null;

            foreach (var rawLine in lines)
            {
                string line = rawLine.Trim();
                if (line.StartsWith("{"))
                {
                    current = new SequenceNumberDefinition();
                }
                else if (line.StartsWith("}") && current != null)
                {
                    if (!string.IsNullOrWhiteSpace(current.Name))
                    {
                        list.Add(current);
                    }
                    current = null;
                }
                else if (current != null && line.Contains(":"))
                {
                    int idx = line.IndexOf(':');
                    string key = line.Substring(0, idx).Trim('\"', ' ', '\t');
                    string val = line.Substring(idx + 1).Trim().TrimEnd(',');

                    switch (key)
                    {
                        case "Id": current.Id = val.Trim('\"'); break;
                        case "Name": current.Name = val.Trim('\"'); break;
                        case "StartNumber": if (long.TryParse(val, out long sn)) current.StartNumber = sn; break;
                        case "EndNumber": if (long.TryParse(val, out long en)) current.EndNumber = en; else current.EndNumber = null; break;
                        case "Increment": if (long.TryParse(val, out long inc)) current.Increment = inc; break;
                        case "Digits": if (int.TryParse(val, out int dig)) current.Digits = dig; break;
                        case "Prefix": current.Prefix = val.Trim('\"'); break;
                        case "Suffix": current.Suffix = val.Trim('\"'); break;
                        case "RepeatCount": if (int.TryParse(val, out int rc)) current.RepeatCount = rc; break;
                        case "CurrentNumber": if (long.TryParse(val, out long cn)) current.CurrentNumber = cn; break;
                        case "IsBuiltIn": current.IsBuiltIn = bool.TryParse(val, out bool bi) && bi; break;
                    }
                }
            }

            return list.Count > 0 ? list : null;
        }

        private static string Escape(string? s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s!.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");
        }
    }
}
