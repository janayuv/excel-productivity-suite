using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelProductivitySuite.Engine.Formulas
{
    public enum FormulaCategory
    {
        All,
        LookupAndReference,
        Financial,
        MathAndStatistics,
        TextManipulation,
        DateAndTime,
        LogicalAndConditional
    }

    public class FormulaParameter
    {
        public string Key { get; }
        public string DisplayName { get; }
        public string DefaultValue { get; }
        public string Description { get; }
        public string Example { get; }

        public FormulaParameter(string key, string displayName, string defaultValue, string description, string example = "")
        {
            Key = key;
            DisplayName = displayName;
            DefaultValue = defaultValue;
            Description = description;
            Example = example;
        }
    }

    public class FormulaTemplate
    {
        public string Id { get; }
        public string Name { get; }
        public FormulaCategory Category { get; }
        public string Description { get; }
        public string TemplatePattern { get; }
        public IReadOnlyList<FormulaParameter> Parameters { get; }

        public FormulaTemplate(
            string id,
            string name,
            FormulaCategory category,
            string description,
            string templatePattern,
            IReadOnlyList<FormulaParameter> parameters)
        {
            Id = id;
            Name = name;
            Category = category;
            Description = description;
            TemplatePattern = templatePattern;
            Parameters = parameters ?? new List<FormulaParameter>();
        }

        public string RenderFormula(IDictionary<string, string>? values)
        {
            string result = TemplatePattern;
            foreach (var p in Parameters)
            {
                string val = p.DefaultValue;
                if (values != null && values.TryGetValue(p.Key, out string? customVal) && !string.IsNullOrWhiteSpace(customVal))
                {
                    val = customVal.Trim();
                }
                result = result.Replace("{" + p.Key + "}", val);
            }

            if (!result.StartsWith("=")) result = "=" + result;
            return result;
        }
    }

    public static class FormulaTemplateLibrary
    {
        private static readonly List<FormulaTemplate> _templates = new List<FormulaTemplate>
        {
            // --- Lookup & Reference ---
            new FormulaTemplate(
                "lookup_2way",
                "Two-Way Matrix Lookup (INDEX + 2 MATCH)",
                FormulaCategory.LookupAndReference,
                "Looks up a value at the intersection of a specific row and column header in a 2D matrix.",
                "=INDEX({MatrixRange}, MATCH({RowLookupValue}, {RowHeadersRange}, 0), MATCH({ColLookupValue}, {ColHeadersRange}, 0))",
                new[]
                {
                    new FormulaParameter("MatrixRange", "Matrix Data Range", "B2:G50", "The rectangular grid containing the values to return"),
                    new FormulaParameter("RowLookupValue", "Row Lookup Value", "A2", "The criteria to match against row headers"),
                    new FormulaParameter("RowHeadersRange", "Row Headers Range", "A2:A50", "The column containing row keys"),
                    new FormulaParameter("ColLookupValue", "Column Lookup Value", "B1", "The criteria to match against column headers"),
                    new FormulaParameter("ColHeadersRange", "Column Headers Range", "B1:G1", "The row containing column keys")
                }),

            new FormulaTemplate(
                "lookup_vlookup_safe",
                "Safe VLOOKUP with Fallback (IFERROR)",
                FormulaCategory.LookupAndReference,
                "Performs a vertical table lookup, safely returning a default fallback value if #N/A is encountered.",
                "=IFERROR(VLOOKUP({LookupValue}, {TableRange}, {ColIndex}, FALSE), {FallbackValue})",
                new[]
                {
                    new FormulaParameter("LookupValue", "Lookup Value", "A2", "The key to look up in the first column"),
                    new FormulaParameter("TableRange", "Table Range", "$D$2:$G$100", "The reference table array"),
                    new FormulaParameter("ColIndex", "Column Index (1-based)", "2", "The column position within TableRange to retrieve"),
                    new FormulaParameter("FallbackValue", "Fallback Default Value", "\"Not Found\"", "Value to show if no match exists")
                }),

            new FormulaTemplate(
                "lookup_left_index_match",
                "Reverse / Left Lookup (INDEX + MATCH)",
                FormulaCategory.LookupAndReference,
                "Retrieves a value from a column to the left of the lookup column, overcoming standard VLOOKUP limitations.",
                "=INDEX({ReturnColumn}, MATCH({LookupValue}, {LookupColumn}, 0))",
                new[]
                {
                    new FormulaParameter("ReturnColumn", "Return Column Range", "$A$2:$A$100", "The column containing the result you want"),
                    new FormulaParameter("LookupValue", "Lookup Value", "D2", "The value to find"),
                    new FormulaParameter("LookupColumn", "Lookup Column Range", "$C$2:$C$100", "The column to search for the match")
                }),

            // --- Financial & Business ---
            new FormulaTemplate(
                "fin_cagr",
                "Compound Annual Growth Rate (CAGR)",
                FormulaCategory.Financial,
                "Calculates the geometric mean annual growth rate of an investment over multiple time periods.",
                "=(({EndValue}/{StartValue})^(1/{Periods}))-1",
                new[]
                {
                    new FormulaParameter("EndValue", "Ending Value Cell", "B10", "The ending value of the investment"),
                    new FormulaParameter("StartValue", "Starting Value Cell", "B2", "The beginning value of the investment"),
                    new FormulaParameter("Periods", "Number of Years / Periods", "5", "The number of compounding years or periods")
                }),

            new FormulaTemplate(
                "fin_pmt",
                "Monthly Loan Payment (PMT)",
                FormulaCategory.Financial,
                "Calculates the fixed monthly installment for a loan based on constant payments and annual interest rate.",
                "=PMT({AnnualRate}/12, {Years}*12, -{Principal})",
                new[]
                {
                    new FormulaParameter("AnnualRate", "Annual Interest Rate", "B2", "Cell containing annual interest rate (e.g. 5%)"),
                    new FormulaParameter("Years", "Loan Term in Years", "B3", "Duration of the loan in years"),
                    new FormulaParameter("Principal", "Loan Principal Amount", "B1", "Initial loan balance borrowed")
                }),

            new FormulaTemplate(
                "fin_variance_pct",
                "Percentage Variance / Growth",
                FormulaCategory.Financial,
                "Computes percentage variance or budget difference between actual performance and target baseline.",
                "=({Actual}-{Target})/{Target}",
                new[]
                {
                    new FormulaParameter("Actual", "Actual Result Cell", "B2", "The realized or actual revenue/cost"),
                    new FormulaParameter("Target", "Target / Budget Cell", "C2", "The budgeted or baseline target")
                }),

            // --- Math & Statistics ---
            new FormulaTemplate(
                "math_weighted_avg",
                "Weighted Average (SUMPRODUCT / SUM)",
                FormulaCategory.MathAndStatistics,
                "Computes the weighted mean of a data series according to designated weighting factors.",
                "=SUMPRODUCT({ValuesRange}, {WeightsRange})/SUM({WeightsRange})",
                new[]
                {
                    new FormulaParameter("ValuesRange", "Values Range", "B2:B20", "The numbers to average"),
                    new FormulaParameter("WeightsRange", "Weights Range", "C2:C20", "The corresponding weights for each number")
                }),

            new FormulaTemplate(
                "math_running_total",
                "Cumulative Running Total",
                FormulaCategory.MathAndStatistics,
                "Creates an expanding running sum anchored at the first row and expanding dynamically down rows.",
                "=SUM({FirstCell}:{ActiveCell})",
                new[]
                {
                    new FormulaParameter("FirstCell", "Anchored First Cell", "$B$2", "Absolute reference to starting cell"),
                    new FormulaParameter("ActiveCell", "Current Relative Cell", "B2", "Relative reference to current row cell")
                }),

            new FormulaTemplate(
                "math_z_score",
                "Statistical Z-Score (Standardized Value)",
                FormulaCategory.MathAndStatistics,
                "Calculates how many standard deviations a value is above or below the population mean.",
                "=({ValueCell}-AVERAGE({DataRange}))/STDEV.S({DataRange})",
                new[]
                {
                    new FormulaParameter("ValueCell", "Value to Standardize", "B2", "The cell value to evaluate"),
                    new FormulaParameter("DataRange", "Dataset Range", "$B$2:$B$100", "The sample distribution range")
                }),

            // --- Text Manipulation ---
            new FormulaTemplate(
                "text_first_word",
                "Extract First Word",
                FormulaCategory.TextManipulation,
                "Extracts the initial word from a space-separated text string.",
                "=LEFT({TextCell}, FIND(\" \", {TextCell}&\" \")-1)",
                new[]
                {
                    new FormulaParameter("TextCell", "Source Text Cell", "A2", "The cell containing full text or full name")
                }),

            new FormulaTemplate(
                "text_last_word",
                "Extract Last Word",
                FormulaCategory.TextManipulation,
                "Extracts the trailing or final word from a sentence or full name.",
                "=TRIM(RIGHT(SUBSTITUTE({TextCell}, \" \", REPT(\" \", 100)), 100))",
                new[]
                {
                    new FormulaParameter("TextCell", "Source Text Cell", "A2", "The cell containing text")
                }),

            new FormulaTemplate(
                "text_domain",
                "Extract Domain from Email Address",
                FormulaCategory.TextManipulation,
                "Isolates the domain name portion after the '@' symbol from an email address.",
                "=MID({EmailCell}, FIND(\"@\", {EmailCell})+1, LEN({EmailCell}))",
                new[]
                {
                    new FormulaParameter("EmailCell", "Email Address Cell", "A2", "The cell containing an email address")
                }),

            // --- Date & Time ---
            new FormulaTemplate(
                "date_workdays",
                "Working Days Between Dates (NETWORKDAYS)",
                FormulaCategory.DateAndTime,
                "Calculates total business working days between two dates, excluding weekends and public holidays.",
                "=NETWORKDAYS({StartDate}, {EndDate}, {HolidaysRange})",
                new[]
                {
                    new FormulaParameter("StartDate", "Start Date Cell", "A2", "Beginning calendar date"),
                    new FormulaParameter("EndDate", "End Date Cell", "B2", "Conclusion calendar date"),
                    new FormulaParameter("HolidaysRange", "Holidays Range (Optional)", "$H$2:$H$15", "Optional range of scheduled holiday dates")
                }),

            new FormulaTemplate(
                "date_age_years",
                "Age in Completed Years (DATEDIF)",
                FormulaCategory.DateAndTime,
                "Calculates completed chronological age or tenure in full elapsed years.",
                "=DATEDIF({BirthDateCell}, TODAY(), \"Y\")",
                new[]
                {
                    new FormulaParameter("BirthDateCell", "Birth Date / Start Date", "A2", "The cell containing birth or start date")
                }),

            new FormulaTemplate(
                "date_eomonth",
                "Last Day of Month (EOMONTH)",
                FormulaCategory.DateAndTime,
                "Returns the serial number for the last day of the month before or after a specified number of months.",
                "=EOMONTH({DateCell}, {MonthsOffset})",
                new[]
                {
                    new FormulaParameter("DateCell", "Base Date Cell", "A2", "The reference date cell"),
                    new FormulaParameter("MonthsOffset", "Months Offset (0 = current)", "0", "0 for current month-end, 1 for next month-end")
                }),

            // --- Logical & Conditional ---
            new FormulaTemplate(
                "logic_band_tier",
                "Tiered Volume Commission / Rating (Nested IF)",
                FormulaCategory.LogicalAndConditional,
                "Assigns categorical performance ratings or commission tiers based on numeric score thresholds.",
                "=IF({ValueCell}>=90, \"Tier A\", IF({ValueCell}>=75, \"Tier B\", IF({ValueCell}>=50, \"Tier C\", \"Tier D\")))",
                new[]
                {
                    new FormulaParameter("ValueCell", "Score / Sales Cell", "B2", "The numeric value to classify")
                })
        };

        public static IReadOnlyList<FormulaTemplate> GetAllTemplates() => _templates;

        public static IReadOnlyList<FormulaTemplate> GetByCategory(FormulaCategory category)
        {
            if (category == FormulaCategory.All) return _templates;
            return _templates.Where(t => t.Category == category).ToList();
        }

        public static IReadOnlyList<FormulaTemplate> Search(string query, FormulaCategory category = FormulaCategory.All)
        {
            var source = category == FormulaCategory.All ? _templates : _templates.Where(t => t.Category == category);
            if (string.IsNullOrWhiteSpace(query)) return source.ToList();

            string q = query.Trim().ToLowerInvariant();
            return source.Where(t =>
                t.Name.ToLowerInvariant().Contains(q) ||
                t.Description.ToLowerInvariant().Contains(q) ||
                t.TemplatePattern.ToLowerInvariant().Contains(q)).ToList();
        }
    }
}
