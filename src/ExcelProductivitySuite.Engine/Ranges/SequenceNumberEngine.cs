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
        /// Gets the standard built-in sequence presets covering numeric, alphabetic, roman, hex, cycles, and dynamic date codes.
        /// </summary>
        public static List<SequenceNumberDefinition> GetDefaultPresets()
        {
            return new List<SequenceNumberDefinition>
            {
                new SequenceNumberDefinition("Simple Numbers (1, 2, 3...)", 1, 1, 1, "", "", null, 1, true, SequenceType.Numeric),
                new SequenceNumberDefinition("Leading Zeros (0001, 0002...)", 1, 1, 4, "", "", null, 1, true, SequenceType.Numeric),
                new SequenceNumberDefinition("Invoice Code (INV-{YYYY}-0001)", 1, 1, 4, "INV-{YYYY}-", "", null, 1, true, SequenceType.Numeric),
                new SequenceNumberDefinition("Monthly Order (ORD-{YYYY}{MM}-0001)", 1, 1, 4, "ORD-{YYYY}{MM}-", "", null, 1, true, SequenceType.Numeric),
                new SequenceNumberDefinition("Employee ID (EMP-001)", 1, 1, 3, "EMP-", "", null, 1, true, SequenceType.Numeric),
                new SequenceNumberDefinition("Uppercase Letters (A, B, C...)", 1, 1, 1, "", "", null, 1, true, SequenceType.AlphaUpper),
                new SequenceNumberDefinition("Lowercase Letters (a, b, c...)", 1, 1, 1, "", "", null, 1, true, SequenceType.AlphaLower),
                new SequenceNumberDefinition("Revision Code (REV-A, REV-B...)", 1, 1, 1, "REV-", "", null, 1, true, SequenceType.AlphaUpper),
                new SequenceNumberDefinition("Roman Numerals (I, II, III...)", 1, 1, 1, "", "", null, 1, true, SequenceType.RomanUpper),
                new SequenceNumberDefinition("Lowercase Roman (i, ii, iii...)", 1, 1, 1, "", "", null, 1, true, SequenceType.RomanLower),
                new SequenceNumberDefinition("Fiscal Quarter Cycle (Q1, Q2, Q3, Q4)", 1, 1, 1, "", "", 4, 1, true, SequenceType.CustomList,
                    new List<string> { "Q1", "Q2", "Q3", "Q4" }),
                new SequenceNumberDefinition("Weekday Cycle (Mon, Tue, Wed...)", 1, 1, 1, "", "", 5, 1, true, SequenceType.CustomList,
                    new List<string> { "Mon", "Tue", "Wed", "Thu", "Fri" }),
                new SequenceNumberDefinition("Hexadecimal Code (0x0001, 0x0002...)", 1, 1, 4, "0x", "", null, 1, true, SequenceType.Hexadecimal),
                new SequenceNumberDefinition("Repeating Cycle (1 to 5)", 1, 1, 1, "", "", 5, 1, true, SequenceType.Numeric),
                new SequenceNumberDefinition("Repeating Pairs (1, 1, 2, 2...)", 1, 1, 1, "", "", null, 2, true, SequenceType.Numeric),
                new SequenceNumberDefinition("Item Suffix (Item-01-A)", 1, 1, 2, "Item-", "-A", null, 1, true, SequenceType.Numeric)
            };
        }

        /// <summary>
        /// Generates a list of formatted sequential strings of the specified length.
        /// </summary>
        public static List<string> GenerateSequence(SequenceNumberDefinition def, int count, out long nextNumber)
        {
            return GenerateSequence(def, count, false, null, null, out nextNumber);
        }

        /// <summary>
        /// Generates a list of formatted sequential strings of the specified length starting from current or start number.
        /// </summary>
        public static List<string> GenerateSequence(SequenceNumberDefinition def, int count, bool continueFromCurrent, out long nextNumber)
        {
            return GenerateSequence(def, count, continueFromCurrent, null, null, out nextNumber);
        }

        /// <summary>
        /// Generates a list of formatted sequential strings with optional reference date and sheet name for dynamic tokens.
        /// </summary>
        public static List<string> GenerateSequence(
            SequenceNumberDefinition def,
            int count,
            bool continueFromCurrent,
            DateTime? referenceDate,
            string? sheetName,
            out long nextNumber)
        {
            if (def == null) throw new ArgumentNullException(nameof(def));
            count = Math.Max(0, count);

            var result = new List<string>(count);
            long val = continueFromCurrent ? def.CurrentNumber : def.StartNumber;
            int rep = 0;

            for (int i = 0; i < count; i++)
            {
                result.Add(def.FormatValue(val, referenceDate, sheetName));
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
        /// Populates a 2D matrix with sequence numbers based on traversal order, filtering, grouping, and randomization options.
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

            object?[,] output = (object?[,])matrix.Clone();

            // 1. Determine cell traversal coordinates in desired order
            var coords = new List<(int r, int c)>();

            switch (options.FillOrder)
            {
                case SequenceFillOrder.Vertical:
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        for (int r = rStart; r <= rEnd; r++)
                        {
                            coords.Add((r, c));
                        }
                    }
                    break;

                case SequenceFillOrder.Horizontal:
                    for (int r = rStart; r <= rEnd; r++)
                    {
                        for (int c = cStart; c <= cEnd; c++)
                        {
                            coords.Add((r, c));
                        }
                    }
                    break;

                case SequenceFillOrder.SerpentineVertical:
                    for (int c = cStart; c <= cEnd; c++)
                    {
                        int colIdx = c - cStart;
                        if (colIdx % 2 == 0)
                        {
                            for (int r = rStart; r <= rEnd; r++)
                                coords.Add((r, c));
                        }
                        else
                        {
                            for (int r = rEnd; r >= rStart; r--)
                                coords.Add((r, c));
                        }
                    }
                    break;

                case SequenceFillOrder.SerpentineHorizontal:
                    for (int r = rStart; r <= rEnd; r++)
                    {
                        int rowIdx = r - rStart;
                        if (rowIdx % 2 == 0)
                        {
                            for (int c = cStart; c <= cEnd; c++)
                                coords.Add((r, c));
                        }
                        else
                        {
                            for (int c = cEnd; c >= cStart; c--)
                                coords.Add((r, c));
                        }
                    }
                    break;
            }

            // 2. Filter eligible coordinates based on StepRows and Blank-only criteria
            var eligibleCoords = new List<(int r, int c)>();
            int skipped = 0;

            foreach (var (r, c) in coords)
            {
                if (options.StepRows > 1 && ((r - rStart) % options.StepRows) != 0)
                {
                    skipped++;
                    continue;
                }

                object? existing = output[r, c];
                bool isBlank = existing == null || string.IsNullOrWhiteSpace(existing.ToString());

                if (options.OnlyFillBlankCells && !isBlank)
                {
                    skipped++;
                    continue;
                }

                eligibleCoords.Add((r, c));
            }

            if (eligibleCoords.Count == 0)
            {
                updatedCount = 0;
                skippedCount = skipped;
                nextNumber = options.ContinueFromLast ? def.CurrentNumber : def.StartNumber;
                return output;
            }

            // 3. Process insertion
            long val = options.ContinueFromLast ? def.CurrentNumber : def.StartNumber;
            int rep = 0;
            int updated = 0;

            // Scenario A: Randomize order within the target cells
            if (options.RandomizeOrder && !options.RestartOnValueChange)
            {
                var valuesToInsert = new List<object?>();
                long genVal = val;
                int genRep = 0;

                for (int i = 0; i < eligibleCoords.Count; i++)
                {
                    string formatted = def.FormatValue(genVal, options.ReferenceDate, options.WorksheetName);
                    if (def.OutputAsNumeric && def.Type == SequenceType.Numeric && long.TryParse(formatted, out long num))
                    {
                        valuesToInsert.Add(num);
                    }
                    else
                    {
                        valuesToInsert.Add(formatted);
                    }

                    genRep++;
                    if (genRep >= def.RepeatCount)
                    {
                        genRep = 0;
                        genVal += def.Increment;
                        if (def.EndNumber.HasValue)
                        {
                            if (def.Increment > 0 && genVal > def.EndNumber.Value) genVal = def.StartNumber;
                            else if (def.Increment < 0 && genVal < def.EndNumber.Value) genVal = def.StartNumber;
                        }
                    }
                }

                // Fisher-Yates shuffle
                var rng = new Random();
                for (int i = valuesToInsert.Count - 1; i > 0; i--)
                {
                    int j = rng.Next(i + 1);
                    var temp = valuesToInsert[i];
                    valuesToInsert[i] = valuesToInsert[j];
                    valuesToInsert[j] = temp;
                }

                for (int i = 0; i < eligibleCoords.Count; i++)
                {
                    var (r, c) = eligibleCoords[i];
                    output[r, c] = valuesToInsert[i];
                    updated++;
                }

                nextNumber = genVal;
            }
            else
            {
                // Scenario B: Sequential fill with optional Grouped Restart on value change
                string? prevKeyStr = null;
                bool hasPrevKey = false;

                foreach (var (r, c) in eligibleCoords)
                {
                    if (options.RestartOnValueChange)
                    {
                        int keyCol = c + options.RestartKeyColumnOffset;
                        if (keyCol >= cStart && keyCol <= cEnd)
                        {
                            object? currentKeyVal = matrix[r, keyCol];
                            string currentKeyStr = currentKeyVal?.ToString() ?? string.Empty;

                            if (hasPrevKey && !string.Equals(currentKeyStr, prevKeyStr, StringComparison.OrdinalIgnoreCase))
                            {
                                val = def.StartNumber;
                                rep = 0;
                            }

                            prevKeyStr = currentKeyStr;
                            hasPrevKey = true;
                        }
                    }

                    string formatted = def.FormatValue(val, options.ReferenceDate, options.WorksheetName);

                    if (def.OutputAsNumeric && def.Type == SequenceType.Numeric && long.TryParse(formatted, out long num))
                    {
                        output[r, c] = num;
                    }
                    else
                    {
                        output[r, c] = formatted;
                    }

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

                nextNumber = val;
            }

            updatedCount = updated;
            skippedCount = skipped;
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

        /// <summary>
        /// Exports a collection of sequence definitions to a JSON file.
        /// </summary>
        public static bool ExportToFile(string filePath, IEnumerable<SequenceNumberDefinition> sequences, out string? error)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                error = "File path is required.";
                return false;
            }

            try
            {
                string json = SerializeJson(sequences.ToList());
                File.WriteAllText(filePath, json, Encoding.UTF8);
                error = null;
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Imports sequence definitions from a JSON file.
        /// </summary>
        public static bool ImportFromFile(string filePath, out List<SequenceNumberDefinition> imported, out string? error)
        {
            imported = new List<SequenceNumberDefinition>();

            if (!File.Exists(filePath))
            {
                error = "Selected file does not exist.";
                return false;
            }

            try
            {
                string json = File.ReadAllText(filePath, Encoding.UTF8);
                var list = DeserializeJson(json);
                if (list != null && list.Count > 0)
                {
                    foreach (var item in list)
                    {
                        item.Id = Guid.NewGuid().ToString("N");
                        item.IsBuiltIn = false;
                    }
                    imported = list;
                    error = null;
                    return true;
                }

                error = "No valid sequence definitions found in file.";
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static string SerializeJson(List<SequenceNumberDefinition> list)
        {
            var sb = new StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < list.Count; i++)
            {
                var s = list[i];
                sb.AppendLine("  {");
                sb.AppendLine($"    \"Id\": \"{Escape(s.Id)}\",");
                sb.AppendLine($"    \"Name\": \"{Escape(s.Name)}\",");
                sb.AppendLine($"    \"Type\": \"{s.Type}\",");
                sb.AppendLine($"    \"CustomListItems\": \"{Escape(string.Join("|", s.CustomListItems ?? new List<string>()))}\",");
                sb.AppendLine($"    \"StartNumber\": {s.StartNumber},");
                sb.AppendLine($"    \"EndNumber\": {(s.EndNumber.HasValue ? s.EndNumber.Value.ToString() : "null")},");
                sb.AppendLine($"    \"Increment\": {s.Increment},");
                sb.AppendLine($"    \"Digits\": {s.Digits},");
                sb.AppendLine($"    \"Prefix\": \"{Escape(s.Prefix)}\",");
                sb.AppendLine($"    \"Suffix\": \"{Escape(s.Suffix)}\",");
                sb.AppendLine($"    \"RepeatCount\": {s.RepeatCount},");
                sb.AppendLine($"    \"CurrentNumber\": {s.CurrentNumber},");
                sb.AppendLine($"    \"UseDynamicPlaceholders\": {(s.UseDynamicPlaceholders ? "true" : "false")},");
                sb.AppendLine($"    \"OutputAsNumeric\": {(s.OutputAsNumeric ? "true" : "false")},");
                sb.AppendLine($"    \"IsBuiltIn\": {(s.IsBuiltIn ? "true" : "false")}");
                sb.Append("  }");
                if (i < list.Count - 1) sb.Append(",");
                sb.AppendLine();
            }
            sb.AppendLine("]");
            return sb.ToString();
        }

        public static List<SequenceNumberDefinition>? DeserializeJson(string json)
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
                        case "Type":
                            if (Enum.TryParse<SequenceType>(val.Trim('\"'), true, out var st))
                                current.Type = st;
                            break;
                        case "CustomListItems":
                            string itemsStr = val.Trim('\"');
                            if (!string.IsNullOrEmpty(itemsStr))
                            {
                                current.CustomListItems = itemsStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                            }
                            break;
                        case "StartNumber": if (long.TryParse(val, out long sn)) current.StartNumber = sn; break;
                        case "EndNumber": if (long.TryParse(val, out long en)) current.EndNumber = en; else current.EndNumber = null; break;
                        case "Increment": if (long.TryParse(val, out long inc)) current.Increment = inc; break;
                        case "Digits": if (int.TryParse(val, out int dig)) current.Digits = dig; break;
                        case "Prefix": current.Prefix = val.Trim('\"'); break;
                        case "Suffix": current.Suffix = val.Trim('\"'); break;
                        case "RepeatCount": if (int.TryParse(val, out int rc)) current.RepeatCount = rc; break;
                        case "CurrentNumber": if (long.TryParse(val, out long cn)) current.CurrentNumber = cn; break;
                        case "UseDynamicPlaceholders": current.UseDynamicPlaceholders = bool.TryParse(val, out bool udp) && udp; break;
                        case "OutputAsNumeric": current.OutputAsNumeric = bool.TryParse(val, out bool oan) && oan; break;
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
