using System;
using ExcelProductivitySuite.Engine.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelProductivitySuite.UnitTests
{
    [TestClass]
    public class TextOperationsTests
    {
        [TestMethod]
        public void Extract_BeforeDelimiter_FirstOccurrence_Succeeds()
        {
            var options = new TextExtractionOptions
            {
                Mode = ExtractionMode.BeforeDelimiter,
                PrimaryDelimiter = ",",
                Occurrence = DelimiterOccurrence.First
            };

            string result = TextExtractor.Extract("Smith, John, Jr.", options);
            Assert.AreEqual("Smith", result);
        }

        [TestMethod]
        public void Extract_BeforeDelimiter_LastOccurrence_Succeeds()
        {
            var options = new TextExtractionOptions
            {
                Mode = ExtractionMode.BeforeDelimiter,
                PrimaryDelimiter = ",",
                Occurrence = DelimiterOccurrence.Last
            };

            string result = TextExtractor.Extract("Smith, John, Jr.", options);
            Assert.AreEqual("Smith, John", result);
        }

        [TestMethod]
        public void Extract_AfterDelimiter_FirstOccurrence_Succeeds()
        {
            var options = new TextExtractionOptions
            {
                Mode = ExtractionMode.AfterDelimiter,
                PrimaryDelimiter = "@",
                Occurrence = DelimiterOccurrence.First
            };

            string result = TextExtractor.Extract("user@domain.com", options);
            Assert.AreEqual("domain.com", result);
        }

        [TestMethod]
        public void Extract_BetweenDelimiters_Succeeds()
        {
            var options = new TextExtractionOptions
            {
                Mode = ExtractionMode.BetweenDelimiters,
                PrimaryDelimiter = "[",
                SecondaryDelimiter = "]"
            };

            string result = TextExtractor.Extract("Order [ORD-9842] Received", options);
            Assert.AreEqual("ORD-9842", result);
        }

        [TestMethod]
        public void Split_ByComma_TrimTokens_Succeeds()
        {
            var options = new TextSplitOptions
            {
                Delimiter = ",",
                TrimTokens = true,
                RemoveEmptyTokens = true
            };

            string[] tokens = TextSplitter.Split("Apple, Banana , Orange, , Grape ", options);
            Assert.AreEqual(4, tokens.Length);
            Assert.AreEqual("Apple", tokens[0]);
            Assert.AreEqual("Banana", tokens[1]);
            Assert.AreEqual("Orange", tokens[2]);
            Assert.AreEqual("Grape", tokens[3]);
        }

        [TestMethod]
        public void RemoveCharacters_Numeric_RemovesDigitsOnly()
        {
            var options = new CharacterRemovalOptions { RemoveNumeric = true };
            string result = CharacterRemover.RemoveCharacters("Suite 2026 Build 42", options);
            Assert.AreEqual("Suite  Build ", result);
        }

        [TestMethod]
        public void RemoveCharacters_NonAlphanumeric_PreservesLettersAndDigits()
        {
            var options = new CharacterRemovalOptions { RemoveNonAlphanumeric = true };
            string result = CharacterRemover.RemoveCharacters("Item #123 (Discount: 15%!)", options);
            Assert.AreEqual("Item 123 Discount 15", result);
        }

        [TestMethod]
        public void Whitespace_CollapseExcess_NormalizesSpaces()
        {
            var options = new WhitespaceOptions { Mode = WhitespaceTrimMode.CollapseExcessInternalSpaces };
            string result = WhitespaceNormalizer.Normalize("   Hello    World   Excel   ", options);
            Assert.AreEqual("Hello World Excel", result);
        }

        [TestMethod]
        public void Whitespace_ConvertNonBreakingSpaces_ReplacesNBSP()
        {
            var options = new WhitespaceOptions { ConvertNonBreakingSpaces = true, Mode = WhitespaceTrimMode.CollapseExcessInternalSpaces };
            string result = WhitespaceNormalizer.Normalize("Price:\u00A0$100.00", options);
            Assert.AreEqual("Price: $100.00", result);
        }

        [TestMethod]
        public void CaseConverter_ProperCase_CapitalizesWords()
        {
            string result = CaseConverter.ConvertCase("microSOFT exCEL proDUctivity", TextCaseMode.ProperCase);
            Assert.AreEqual("Microsoft Excel Productivity", result);
        }

        [TestMethod]
        public void CaseConverter_SentenceCase_CapitalizesSentences()
        {
            string result = CaseConverter.ConvertCase("first sentence. second sentence! third?", TextCaseMode.SentenceCase);
            Assert.AreEqual("First sentence. Second sentence! Third?", result);
        }

        [TestMethod]
        public void CaseConverter_ToggleCase_InvertsCase()
        {
            string result = CaseConverter.ConvertCase("Excel 2016 PRO", TextCaseMode.ToggleCase);
            Assert.AreEqual("eXCEL 2016 pro", result);
        }

        [TestMethod]
        public void AffixAppender_AddsPrefixAndSuffix_RespectsConditions()
        {
            var options = new AffixOptions
            {
                Prefix = "ID-",
                Suffix = "-PROD",
                SkipIfAlreadyContains = true
            };

            string result1 = AffixAppender.ApplyAffixes("1042", options);
            Assert.AreEqual("ID-1042-PROD", result1);

            // If already contains prefix, does not duplicate
            string result2 = AffixAppender.ApplyAffixes("ID-5500", options);
            Assert.AreEqual("ID-5500-PROD", result2);
        }

        [TestMethod]
        public void PositionInserter_FromStart_InsertsAtTargetIndex()
        {
            var options = new PositionInsertOptions
            {
                TextToInsert = "-",
                PositionIndex = 4, // After 3rd character
                FromEnd = false
            };

            string result = PositionInserter.Insert("ABCDEF", options);
            Assert.AreEqual("ABC-DEF", result);
        }
    }
}
