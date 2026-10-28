using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Formulas
{
    public class FormulaReplaceOptions
    {
        public string FindText { get; set; } = string.Empty;
        public string ReplaceWith { get; set; } = string.Empty;
        public bool MatchCase { get; set; } = false;
        public bool MatchWholeWord { get; set; } = false;
        public bool IgnoreStringLiterals { get; set; } = true;

        public FormulaReplaceOptions(string findText, string replaceWith)
        {
            FindText = findText ?? string.Empty;
            ReplaceWith = replaceWith ?? string.Empty;
        }

        public FormulaReplaceOptions() { }
    }

    public class FormulaChangeItem
    {
        public int Row { get; }
        public int Column { get; }
        public string CellAddress { get; }
        public string OriginalFormula { get; }
        public string ModifiedFormula { get; }
        public int ReplacementsCount { get; }

        public FormulaChangeItem(int row, int col, string address, string originalFormula, string modifiedFormula, int replacementsCount)
        {
            Row = row;
            Column = col;
            CellAddress = address;
            OriginalFormula = originalFormula;
            ModifiedFormula = modifiedFormula;
            ReplacementsCount = replacementsCount;
        }
    }

    public class FormulaBatchReplaceResult
    {
        public object?[,] UpdatedFormulas { get; }
        public int TotalCellsScanned { get; }
        public int FormulasFound { get; }
        public int FormulasModified { get; }
        public int TotalReplacements { get; }
        public IReadOnlyList<FormulaChangeItem> Changes { get; }

        public FormulaBatchReplaceResult(
            object?[,] updatedFormulas,
            int totalCellsScanned,
            int formulasFound,
            int formulasModified,
            int totalReplacements,
            IReadOnlyList<FormulaChangeItem> changes)
        {
            UpdatedFormulas = updatedFormulas;
            TotalCellsScanned = totalCellsScanned;
            FormulasFound = formulasFound;
            FormulasModified = formulasModified;
            TotalReplacements = totalReplacements;
            Changes = changes;
        }
    }

    /// <summary>
    /// High-performance engine for batch finding and replacing sheet references, range tokens,
    /// or function names in Excel formulas while respecting string literals and syntax rules.
    /// </summary>
    public static class FormulaBatchReplacer
    {
        private static readonly Regex StringLiteralRegex = new Regex(@"""(?:[^""]|"""")*""", RegexOptions.Compiled);

        /// <summary>
        /// Replaces target text inside a single formula string.
        /// </summary>
        public static string ReplaceInFormula(string formula, FormulaReplaceOptions options, out int replacementsCount)
        {
            replacementsCount = 0;
            if (string.IsNullOrWhiteSpace(formula) || string.IsNullOrEmpty(options.FindText))
            {
                return formula;
            }

            bool hasEquals = formula.StartsWith("=");
            string expression = hasEquals ? formula.Substring(1) : formula;

            if (options.IgnoreStringLiterals)
            {
                var sb = new StringBuilder();
                if (hasEquals) sb.Append('=');

                int lastIndex = 0;
                var matches = StringLiteralRegex.Matches(expression);

                foreach (Match m in matches)
                {
                    if (m.Index > lastIndex)
                    {
                        string codePart = expression.Substring(lastIndex, m.Index - lastIndex);
                        string replacedCode = ReplaceCodeSegment(codePart, options, out int partCount);
                        sb.Append(replacedCode);
                        replacementsCount += partCount;
                    }

                    // Keep string literal verbatim
                    sb.Append(m.Value);
                    lastIndex = m.Index + m.Length;
                }

                if (lastIndex < expression.Length)
                {
                    string remainingCode = expression.Substring(lastIndex);
                    string replacedCode = ReplaceCodeSegment(remainingCode, options, out int partCount);
                    sb.Append(replacedCode);
                    replacementsCount += partCount;
                }

                return sb.ToString();
            }
            else
            {
                string replaced = ReplaceCodeSegment(expression, options, out replacementsCount);
                return hasEquals ? "=" + replaced : replaced;
            }
        }

        private static string ReplaceCodeSegment(string code, FormulaReplaceOptions options, out int replacementsCount)
        {
            replacementsCount = 0;
            if (string.IsNullOrEmpty(code) || string.IsNullOrEmpty(options.FindText)) return code;

            if (options.MatchWholeWord)
            {
                // Word boundary check that works for cell addresses and alphanumeric tokens
                string pattern = $@"(?<![A-Za-z0-9_]){Regex.Escape(options.FindText)}(?![A-Za-z0-9_])";
                var regexOptions = RegexOptions.None;
                if (!options.MatchCase) regexOptions |= RegexOptions.IgnoreCase;

                var regex = new Regex(pattern, regexOptions);
                var matches = regex.Matches(code);
                replacementsCount = matches.Count;
                return regex.Replace(code, options.ReplaceWith);
            }
            else
            {
                var comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                int count = 0;
                int idx = 0;
                var sb = new StringBuilder();

                while ((idx = code.IndexOf(options.FindText, idx, comparison)) != -1)
                {
                    count++;
                    idx += options.FindText.Length;
                }

                replacementsCount = count;
                if (count == 0) return code;

                // Simple regex replace for case handling
                var regexOptions = RegexOptions.None;
                if (!options.MatchCase) regexOptions |= RegexOptions.IgnoreCase;
                return Regex.Replace(code, Regex.Escape(options.FindText), options.ReplaceWith, regexOptions);
            }
        }

        /// <summary>
        /// Scans and replaces references across a 2D matrix of formulas.
        /// </summary>
        public static FormulaBatchReplaceResult ProcessFormulas(
            object?[,] rawFormulas,
            FormulaReplaceOptions options,
            int startRow = 1,
            int startCol = 1)
        {
            if (rawFormulas == null)
            {
                return new FormulaBatchReplaceResult(new object?[0, 0], 0, 0, 0, 0, new List<FormulaChangeItem>());
            }

            int rStart = rawFormulas.GetLowerBound(0);
            int rEnd = rawFormulas.GetUpperBound(0);
            int cStart = rawFormulas.GetLowerBound(1);
            int cEnd = rawFormulas.GetUpperBound(1);

            int rows = rEnd - rStart + 1;
            int cols = cEnd - cStart + 1;
            var updated = new object?[rows, cols];
            var changes = new List<FormulaChangeItem>();

            int totalCells = 0;
            int formulasFound = 0;
            int formulasModified = 0;
            int totalReplacements = 0;

            for (int r = rStart; r <= rEnd; r++)
            {
                int destR = r - rStart;
                for (int c = cStart; c <= cEnd; c++)
                {
                    int destC = c - cStart;
                    totalCells++;

                    object? cellVal = rawFormulas[r, c];
                    string? formulaStr = cellVal?.ToString();

                    if (!string.IsNullOrWhiteSpace(formulaStr) && formulaStr!.StartsWith("="))
                    {
                        formulasFound++;
                        string modified = ReplaceInFormula(formulaStr, options, out int cellReplacements);

                        if (cellReplacements > 0 && !string.Equals(formulaStr, modified, StringComparison.Ordinal))
                        {
                            formulasModified++;
                            totalReplacements += cellReplacements;

                            int sheetRow = startRow + destR;
                            int sheetCol = startCol + destC;
                            string cellAddr = $"R{sheetRow}C{sheetCol}";

                            changes.Add(new FormulaChangeItem(
                                sheetRow,
                                sheetCol,
                                cellAddr,
                                formulaStr,
                                modified,
                                cellReplacements));

                            updated[destR, destC] = modified;
                        }
                        else
                        {
                            updated[destR, destC] = formulaStr;
                        }
                    }
                    else
                    {
                        updated[destR, destC] = cellVal;
                    }
                }
            }

            return new FormulaBatchReplaceResult(
                updated,
                totalCells,
                formulasFound,
                formulasModified,
                totalReplacements,
                changes);
        }
    }
}
