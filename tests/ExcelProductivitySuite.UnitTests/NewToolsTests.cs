using System;
using System.Collections.Generic;
using System.Linq;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.DataCleaning;
using ExcelProductivitySuite.Engine.DataValidation;
using ExcelProductivitySuite.Engine.Formatting;
using ExcelProductivitySuite.Engine.Text;
using ExcelProductivitySuite.Engine.Worksheets;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class NewToolsTests
    {
        [TestMethod]
        public void SequenceSheetGenerator_ShouldGenerateNumbersWithPadding()
        {
            var options = new SequenceSheetOptions
            {
                Type = SequenceType.Numbers,
                Prefix = "Data_",
                Suffix = "",
                StartNumber = 1,
                Count = 3,
                Step = 1,
                ZeroPadding = 3
            };

            var names = SequenceSheetGenerator.GenerateNames(options);

            Assert.AreEqual(3, names.Count);
            Assert.AreEqual("Data_001", names[0]);
            Assert.AreEqual("Data_002", names[1]);
            Assert.AreEqual("Data_003", names[2]);
        }

        [TestMethod]
        public void SequenceSheetGenerator_ShouldGenerateMonths()
        {
            var options = new SequenceSheetOptions
            {
                Type = SequenceType.Months,
                Prefix = "",
                Suffix = "_Report",
                StartNumber = 1, // Jan
                Count = 3,
                FullName = false
            };

            var names = SequenceSheetGenerator.GenerateNames(options);

            Assert.AreEqual(3, names.Count);
            Assert.AreEqual("Jan_Report", names[0]);
            Assert.AreEqual("Feb_Report", names[1]);
            Assert.AreEqual("Mar_Report", names[2]);
        }

        [TestMethod]
        public void SequenceSheetGenerator_ShouldGenerateQuarters()
        {
            var options = new SequenceSheetOptions
            {
                Type = SequenceType.Quarters,
                Prefix = "",
                Suffix = "",
                StartNumber = 3,
                Count = 3,
                Year = 2026
            };

            var names = SequenceSheetGenerator.GenerateNames(options);

            Assert.AreEqual(3, names.Count);
            Assert.AreEqual("Q3_2026", names[0]);
            Assert.AreEqual("Q4_2026", names[1]);
            Assert.AreEqual("Q1_2027", names[2]);
        }

        [TestMethod]
        public void SequenceSheetGenerator_ShouldValidateAndDetectDuplicates()
        {
            var proposed = new[] { "Sheet1", "ValidSheet", "Invalid/Sheet", "Sheet1" };
            var existing = new[] { "Sheet1" };

            var (valid, errors) = SequenceSheetGenerator.ValidateGeneratedNames(proposed, existing);

            Assert.AreEqual(1, valid.Count);
            Assert.AreEqual("ValidSheet", valid[0]);
            Assert.IsTrue(errors.Count >= 3);
        }

        [TestMethod]
        public void NumberExtractor_ShouldExtractIntegersAndDecimals()
        {
            string text = "Order #1042 has 3 items totaling $159.95 with -5.00 discount";

            var optAll = new NumberExtractionOptions
            {
                Mode = NumberExtractionMode.AllJoined,
                AllowDecimals = true,
                AllowNegative = true,
                Delimiter = ", "
            };
            string all = NumberExtractor.ExtractFromText(text, optAll);
            Assert.AreEqual("1042, 3, 159.95, -5.00", all);

            var optFirst = new NumberExtractionOptions
            {
                Mode = NumberExtractionMode.FirstOnly,
                AllowDecimals = true,
                AllowNegative = true
            };
            string first = NumberExtractor.ExtractFromText(text, optFirst);
            Assert.AreEqual("1042", first);

            var optSum = new NumberExtractionOptions
            {
                Mode = NumberExtractionMode.Sum,
                AllowDecimals = true,
                AllowNegative = true
            };
            string sum = NumberExtractor.ExtractFromText("Item 10 and 20.5 with -5", optSum);
            Assert.AreEqual("25.5", sum);
        }

        [TestMethod]
        public void NumberExtractor_ShouldProcessMatrix()
        {
            var matrix = new object[,]
            {
                { "Batch 45", "Item 10.5" },
                { "No digits", "Qty 99" }
            };

            var options = new NumberExtractionOptions
            {
                Mode = NumberExtractionMode.FirstOnly,
                AllowDecimals = true
            };

            var result = NumberExtractor.ExtractFromMatrix(matrix, options);

            Assert.AreEqual("45", result[0, 0]);
            Assert.AreEqual("10.5", result[0, 1]);
            Assert.IsNull(result[1, 0]);
            Assert.AreEqual("99", result[1, 1]);
        }

        [TestMethod]
        public void EmailExtractor_ShouldExtractAndDeduplicateEmails()
        {
            string text = "Contact alice@example.com or support@company.org. Backup: ALICE@example.com.";

            var optFirst = new EmailExtractionOptions
            {
                Mode = EmailExtractionMode.FirstOnly
            };
            string first = EmailExtractor.ExtractFromText(text, optFirst);
            Assert.AreEqual("alice@example.com", first);

            var optAllDedup = new EmailExtractionOptions
            {
                Mode = EmailExtractionMode.AllJoined,
                Deduplicate = true,
                Delimiter = "; "
            };
            string all = EmailExtractor.ExtractFromText(text, optAllDedup);
            Assert.AreEqual("alice@example.com; support@company.org", all);
        }

        [TestMethod]
        public void EmailExtractor_ShouldProcessMatrix()
        {
            var matrix = new object[,]
            {
                { "Lead: bob@dev.io", "No email here" }
            };

            var opt = new EmailExtractionOptions { Mode = EmailExtractionMode.FirstOnly };
            var result = EmailExtractor.ExtractFromMatrix(matrix, opt);

            Assert.AreEqual("bob@dev.io", result[0, 0]);
            Assert.IsNull(result[0, 1]);
        }

        [TestMethod]
        public void UrlExtractor_ShouldExtractAndNormalizeUrls()
        {
            string text = "Visit our docs at https://docs.example.com/api/v1?page=2 and site www.mysite.org. Also ftp://files.net/pub/data.zip.";

            var optFirst = new UrlExtractionOptions { Mode = UrlExtractionMode.FirstOnly };
            string first = UrlExtractor.ExtractFromText(text, optFirst);
            Assert.AreEqual("https://docs.example.com/api/v1?page=2", first);

            var optAll = new UrlExtractionOptions
            {
                Mode = UrlExtractionMode.AllJoined,
                NormalizeScheme = true,
                Delimiter = " | "
            };
            string all = UrlExtractor.ExtractFromText(text, optAll);
            Assert.AreEqual("https://docs.example.com/api/v1?page=2 | https://www.mysite.org | ftp://files.net/pub/data.zip", all);

            var matrix = new object[,]
            {
                { "Check www.contoso.com.", "No links" }
            };
            var matResult = UrlExtractor.ExtractFromMatrix(matrix, optFirst);
            Assert.AreEqual("https://www.contoso.com", matResult[0, 0]);
            Assert.IsNull(matResult[0, 1]);
        }

        [TestMethod]
        public void BatchReplacer_ShouldExecuteMultipleReplacementRules()
        {
            var rules = new List<BatchReplaceRule>
            {
                new("St.", "Street"),
                new("Ave.", "Avenue"),
                new(@"\b\d{5}\b", "[ZIP]", matchCase: false, matchEntireCell: false, useRegex: true)
            };

            string text = "123 Elm St. and 456 Maple Ave. ZIP 90210";
            string replaced = BatchReplacer.ReplaceInText(text, rules, out int count);

            Assert.AreEqual("123 Elm Street and 456 Maple Avenue ZIP [ZIP]", replaced);
            Assert.AreEqual(3, count);

            var matrix = new object[,]
            {
                { "Elm St.", "Oak Ave." }
            };
            var matResult = BatchReplacer.ReplaceInMatrix(matrix, rules, out int totalCount);
            Assert.AreEqual("Elm Street", matResult[0, 0]);
            Assert.AreEqual("Oak Avenue", matResult[0, 1]);
            Assert.AreEqual(2, totalCount);
        }

        [TestMethod]
        public void FrequencyDistributionCalculator_ShouldComputeFrequenciesAndPercentages()
        {
            var matrix = new object[,]
            {
                { "Red" },
                { "Blue" },
                { "Red" },
                { "Green" },
                { "Red" },
                { "Blue" }
            };

            var options = new FrequencyDistributionOptions
            {
                SortOrder = FrequencySortOrder.FrequencyDescending
            };

            var records = FrequencyDistributionCalculator.Calculate(matrix, options);

            Assert.AreEqual(3, records.Count);
            Assert.AreEqual("Red", records[0].Value);
            Assert.AreEqual(3, records[0].Count);
            Assert.AreEqual(50.0, records[0].Percent);
            Assert.AreEqual(3, records[0].CumulativeCount);
            Assert.AreEqual(50.0, records[0].CumulativePercent);

            Assert.AreEqual("Blue", records[1].Value);
            Assert.AreEqual(2, records[1].Count);

            var outputMatrix = FrequencyDistributionCalculator.GenerateMatrix(records, includeHeaders: true);
            Assert.AreEqual(4, outputMatrix.GetLength(0)); // Header + 3 rows
            Assert.AreEqual(5, outputMatrix.GetLength(1));
            Assert.AreEqual("Value", outputMatrix[0, 0]);
            Assert.AreEqual("Red", outputMatrix[1, 0]);
            Assert.AreEqual(3, outputMatrix[1, 1]);
        }

        [TestMethod]
        public void CellDimensionHelper_ShouldConvertUnitsCorrectly()
        {
            // 1 inch = 72 points
            double ptsFromInches = CellDimensionHelper.ToPoints(1.0, CellDimensionUnit.Inches);
            Assert.AreEqual(72.0, ptsFromInches, 0.001);

            double inchesFromPts = CellDimensionHelper.FromPoints(72.0, CellDimensionUnit.Inches);
            Assert.AreEqual(1.0, inchesFromPts, 0.001);

            // Row height clamp
            Assert.AreEqual(409.5, CellDimensionHelper.ClampRowHeight(500.0));
            Assert.AreEqual(0.0, CellDimensionHelper.ClampRowHeight(-10.0));
            Assert.AreEqual(25.0, CellDimensionHelper.ClampRowHeight(25.0));
        }

        [TestMethod]
        public void DropdownListBuilder_ShouldFormatAndValidateDropdownList()
        {
            string input = "Alpha, Beta, Gamma\nDelta, Alpha,   Epsilon  ";
            string built = DropdownListBuilder.BuildListString(input, deduplicate: true);

            Assert.AreEqual("Alpha,Beta,Gamma,Delta,Epsilon", built);

            bool valid = DropdownListBuilder.ValidateList(built, isFormula: false, out string? errMsg);
            Assert.IsTrue(valid);
            Assert.IsNull(errMsg);

            // Exceeding 255 chars
            string longList = string.Join(",", Enumerable.Repeat("LongItemName1234567890", 20));
            bool longValid = DropdownListBuilder.ValidateList(longList, isFormula: false, out string? longErr);
            Assert.IsFalse(longValid);
            Assert.IsNotNull(longErr);
        }
    }
}
