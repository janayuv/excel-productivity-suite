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

        // =========================================================================
        // Batch 2: Core Text Expansion Unit Tests
        // =========================================================================

        #region TextReverser Tests

        [TestMethod]
        public void TextReverser_ByCharacter_StandardString_InvertsLetters()
        {
            var options = new TextReverseOptions { Mode = TextReverseMode.ByCharacter };
            string result = TextReverser.Reverse("Antigravity", options);
            Assert.AreEqual("ytivargitnA", result);
        }

        [TestMethod]
        public void TextReverser_ByCharacter_PreservesSurrogatePairs()
        {
            var options = new TextReverseOptions { Mode = TextReverseMode.ByCharacter };
            // String with surrogate pair emoji: A 🚀 B
            string input = "A \uD83D\uDE80 B";
            string result = TextReverser.Reverse(input, options);
            Assert.AreEqual("B \uD83D\uDE80 A", result);
        }

        [TestMethod]
        public void TextReverser_ByWord_InvertsWordOrder()
        {
            var options = new TextReverseOptions { Mode = TextReverseMode.ByWord };
            string result = TextReverser.Reverse("The quick brown fox jumps", options);
            Assert.AreEqual("jumps fox brown quick The", result);
        }

        [TestMethod]
        public void TextReverser_ByDelimiter_InvertsTokens()
        {
            var options = new TextReverseOptions
            {
                Mode = TextReverseMode.ByDelimiter,
                Delimiter = ";"
            };
            string result = TextReverser.Reverse("first;second;third;fourth", options);
            Assert.AreEqual("fourth;third;second;first", result);
        }

        [TestMethod]
        public void TextReverser_EmptyOrNull_ReturnsEmpty()
        {
            var options = new TextReverseOptions { Mode = TextReverseMode.ByCharacter };
            Assert.AreEqual("", TextReverser.Reverse("", options));
            Assert.AreEqual("", TextReverser.Reverse(null!, options));
        }

        #endregion

        #region TextCombiner Tests

        [TestMethod]
        public void TextCombiner_CombineList_WithDelimiterAndAffixes()
        {
            var items = new[] { "Alpha", "Beta", "Gamma" };
            var options = new CombineOptions
            {
                Delimiter = " | ",
                Prefix = "[",
                Suffix = "]"
            };

            string result = TextCombiner.CombineList(items, options);
            Assert.AreEqual("[Alpha | Beta | Gamma]", result);
        }

        [TestMethod]
        public void TextCombiner_CombineList_SkipsBlankCells()
        {
            var items = new[] { "North", "", null, "   ", "South" };
            var options = new CombineOptions
            {
                Delimiter = ", ",
                SkipBlankCells = true,
                TrimTokens = true
            };

            string result = TextCombiner.CombineList(items, options);
            Assert.AreEqual("North, South", result);
        }

        [TestMethod]
        public void TextCombiner_Matrix_AcrossColumns()
        {
            var matrix = new object?[2, 3]
            {
                { "John", "M", "Doe" },
                { "Jane", null, "Smith" }
            };

            var options = new CombineOptions
            {
                Direction = CombineDirection.AcrossColumns,
                Delimiter = " "
            };

            var result = TextCombiner.CombineMatrix(matrix, options);
            Assert.AreEqual(2, result.GetLength(0));
            Assert.AreEqual(1, result.GetLength(1));
            Assert.AreEqual("John M Doe", result[0, 0]);
            Assert.AreEqual("Jane Smith", result[1, 0]);
        }

        [TestMethod]
        public void TextCombiner_Matrix_DownRows()
        {
            var matrix = new object?[2, 2]
            {
                { "Item 1", "Red" },
                { "Item 2", "Blue" }
            };

            var options = new CombineOptions
            {
                Direction = CombineDirection.DownRows,
                Delimiter = " / "
            };

            var result = TextCombiner.CombineMatrix(matrix, options);
            Assert.AreEqual(1, result.GetLength(0));
            Assert.AreEqual(2, result.GetLength(1));
            Assert.AreEqual("Item 1 / Item 2", result[0, 0]);
            Assert.AreEqual("Red / Blue", result[0, 1]);
        }

        [TestMethod]
        public void TextCombiner_Matrix_EntireSelection()
        {
            var matrix = new object?[2, 2]
            {
                { "A", "B" },
                { "C", "D" }
            };

            var options = new CombineOptions
            {
                Direction = CombineDirection.EntireSelection,
                Delimiter = "-"
            };

            var result = TextCombiner.CombineMatrix(matrix, options);
            Assert.AreEqual(1, result.GetLength(0));
            Assert.AreEqual(1, result.GetLength(1));
            Assert.AreEqual("A-B-C-D", result[0, 0]);
        }

        #endregion

        #region TextCleaner Tests

        [TestMethod]
        public void TextCleaner_RemovesAsciiControlCodes()
        {
            var options = new TextCleanOptions
            {
                RemoveControlCodes = true,
                PreserveLineBreaks = false,
                PreserveTabs = false
            };

            // Bell (7), Backspace (8), Unit Separator (31), Delete (127)
            string dirty = "Valid\u0007Text\u0008With\u001FControls\u007F!";
            string result = TextCleaner.Clean(dirty, options);
            Assert.AreEqual("ValidTextWithControls!", result);
        }

        [TestMethod]
        public void TextCleaner_RemovesInvisibleZeroWidthChars()
        {
            var options = new TextCleanOptions { RemoveInvisibleZeroWidth = true };
            // Zero-width space (\u200B) and BOM (\uFEFF)
            string dirty = "Hello\u200BWorld\uFEFF!";
            string result = TextCleaner.Clean(dirty, options);
            Assert.AreEqual("HelloWorld!", result);
        }

        [TestMethod]
        public void TextCleaner_DecodesHtmlEntities()
        {
            var options = new TextCleanOptions { DecodeHtmlEntities = true };
            string dirty = "Tom &amp; Jerry &lt;&gt; &quot;Classic&quot;";
            string result = TextCleaner.Clean(dirty, options);
            Assert.AreEqual("Tom & Jerry <> \"Classic\"", result);
        }

        [TestMethod]
        public void TextCleaner_ReplacesLineBreaksAndTabs_WhenPreserveDisabled()
        {
            var options = new TextCleanOptions
            {
                PreserveLineBreaks = false,
                PreserveTabs = false
            };
            string dirty = "Line1\r\nLine2\tLine3";
            string result = TextCleaner.Clean(dirty, options);
            Assert.AreEqual("Line1 Line2 Line3", result);
        }

        #endregion

        #region NumberTextConverter Tests

        [TestMethod]
        public void NumberTextConverter_ParsesStandardNumericString()
        {
            var options = new NumberTextConversionOptions();
            object? result = NumberTextConverter.ConvertToNumber("12345.67", options);
            Assert.IsInstanceOfType(result, typeof(double));
            Assert.AreEqual(12345.67, (double)result!, 0.0001);
        }

        [TestMethod]
        public void NumberTextConverter_StripsCurrencySymbols()
        {
            var options = new NumberTextConversionOptions { AllowCurrencySymbols = true };

            object? usd = NumberTextConverter.ConvertToNumber("$1,234.50", options);
            Assert.AreEqual(1234.50, (double)usd!, 0.0001);

            object? eur = NumberTextConverter.ConvertToNumber("€ 999.00", options);
            Assert.AreEqual(999.0, (double)eur!, 0.0001);

            object? gbp = NumberTextConverter.ConvertToNumber("£45.25", options);
            Assert.AreEqual(45.25, (double)gbp!, 0.0001);

            object? inr = NumberTextConverter.ConvertToNumber("₹ 1,500.75", options);
            Assert.AreEqual(1500.75, (double)inr!, 0.0001);
        }

        [TestMethod]
        public void NumberTextConverter_ParsesAccountingParentheses_AsNegative()
        {
            var options = new NumberTextConversionOptions { AllowParenthesesForNegative = true };
            object? result = NumberTextConverter.ConvertToNumber("(1,250.75)", options);
            Assert.AreEqual(-1250.75, (double)result!, 0.0001);
        }

        [TestMethod]
        public void NumberTextConverter_ParsesPercentages()
        {
            var options = new NumberTextConversionOptions { AllowPercentage = true };
            object? result = NumberTextConverter.ConvertToNumber("25.5%", options);
            Assert.AreEqual(0.255, (double)result!, 0.0001);
        }

        [TestMethod]
        public void NumberTextConverter_SkipsNonNumericWhenConfigured()
        {
            var options = new NumberTextConversionOptions { SkipNonNumeric = true };
            object? result = NumberTextConverter.ConvertToNumber("NotANumber", options);
            Assert.AreEqual("NotANumber", result);
        }

        #endregion

        #region DateTextConverter Tests

        [TestMethod]
        public void DateTextConverter_ParsesIso8601_ToIsoFormat()
        {
            var options = new DateTextConversionOptions { OutputFormat = DateOutputFormat.Iso8601 };
            object? result = DateTextConverter.ConvertToDate("2026-03-15", options);
            Assert.AreEqual("2026-03-15", result);
        }

        [TestMethod]
        public void DateTextConverter_CulturePreference_EU_ParsesDayFirst()
        {
            var options = new DateTextConversionOptions
            {
                CulturePreference = DateCulturePreference.DayMonthYear,
                OutputFormat = DateOutputFormat.Iso8601
            };
            // 05/11/2026 is 5th Nov in EU
            object? result = DateTextConverter.ConvertToDate("05/11/2026", options);
            Assert.AreEqual("2026-11-05", result);
        }

        [TestMethod]
        public void DateTextConverter_CulturePreference_US_ParsesMonthFirst()
        {
            var options = new DateTextConversionOptions
            {
                CulturePreference = DateCulturePreference.MonthDayYear,
                OutputFormat = DateOutputFormat.Iso8601
            };
            // 05/11/2026 is May 11th in US
            object? result = DateTextConverter.ConvertToDate("05/11/2026", options);
            Assert.AreEqual("2026-05-11", result);
        }

        [TestMethod]
        public void DateTextConverter_FormatsAsLongDate()
        {
            var options = new DateTextConversionOptions
            {
                OutputFormat = DateOutputFormat.LongDate
            };
            object? result = DateTextConverter.ConvertToDate("2026-01-25", options);
            Assert.AreEqual("January 25, 2026", result);
        }

        [TestMethod]
        public void DateTextConverter_ReturnsDateTimeStruct()
        {
            var options = new DateTextConversionOptions { OutputFormat = DateOutputFormat.DateTimeObject };
            object? result = DateTextConverter.ConvertToDate("2026-07-04", options);
            Assert.IsInstanceOfType(result, typeof(DateTime));
            var dt = (DateTime)result!;
            Assert.AreEqual(2026, dt.Year);
            Assert.AreEqual(7, dt.Month);
            Assert.AreEqual(4, dt.Day);
        }

        #endregion

        #region NumberSpeller Tests

        [TestMethod]
        public void NumberSpeller_SpellsIntegers_UpToTrillions()
        {
            var options = new NumberSpellerOptions
            {
                Mode = SpellerMode.PureNumber,
                LetterCase = SpellerCase.TitleCase
            };

            Assert.AreEqual("Zero", NumberSpeller.Spell(0, options));
            Assert.AreEqual("Forty-Two", NumberSpeller.Spell(42, options));
            Assert.AreEqual("One Thousand Two Hundred Thirty-Four", NumberSpeller.Spell(1234, options));
            Assert.AreEqual("One Million", NumberSpeller.Spell(1000000, options));
            Assert.AreEqual("One Billion", NumberSpeller.Spell(1000000000, options));
            Assert.AreEqual("One Trillion", NumberSpeller.Spell(1000000000000, options));
        }

        [TestMethod]
        public void NumberSpeller_SpellsUSD_Currency()
        {
            var options = new NumberSpellerOptions
            {
                Mode = SpellerMode.Currency,
                Currency = SpellerCurrency.USD,
                LetterCase = SpellerCase.TitleCase,
                IncludeCentsIfZero = true
            };

            string result1 = NumberSpeller.Spell(125.50, options);
            Assert.AreEqual("One Hundred Twenty-Five Dollars and Fifty Cents", result1);

            string result2 = NumberSpeller.Spell(1.00, options);
            Assert.AreEqual("One Dollar and Zero Cents", result2);
        }

        [TestMethod]
        public void NumberSpeller_SpellsEUR_Currency()
        {
            var options = new NumberSpellerOptions
            {
                Mode = SpellerMode.Currency,
                Currency = SpellerCurrency.EUR,
                LetterCase = SpellerCase.TitleCase
            };

            string result = NumberSpeller.Spell(500.75, options);
            Assert.AreEqual("Five Hundred Euros and Seventy-Five Cents", result);
        }

        [TestMethod]
        public void NumberSpeller_SpellsINR_Currency()
        {
            var options = new NumberSpellerOptions
            {
                Mode = SpellerMode.Currency,
                Currency = SpellerCurrency.INR,
                LetterCase = SpellerCase.TitleCase
            };

            string result = NumberSpeller.Spell(1500.50, options);
            Assert.AreEqual("One Thousand Five Hundred Rupees and Fifty Paise", result);
        }

        [TestMethod]
        public void NumberSpeller_HandlesNegativeValues()
        {
            var options = new NumberSpellerOptions
            {
                Mode = SpellerMode.PureNumber,
                LetterCase = SpellerCase.TitleCase
            };

            string result = NumberSpeller.Spell(-45.2, options);
            Assert.IsTrue(result.StartsWith("Negative Forty-Five"));
        }

        [TestMethod]
        public void NumberSpeller_SupportsCasingOptions()
        {
            var upperOpts = new NumberSpellerOptions { Mode = SpellerMode.PureNumber, LetterCase = SpellerCase.UpperCase };
            Assert.AreEqual("FIFTY", NumberSpeller.Spell(50, upperOpts));

            var lowerOpts = new NumberSpellerOptions { Mode = SpellerMode.PureNumber, LetterCase = SpellerCase.LowerCase };
            Assert.AreEqual("fifty", NumberSpeller.Spell(50, lowerOpts));
        }

        #endregion

        #region PositionRemover Tests

        [TestMethod]
        public void PositionRemover_FromLeft_RemovesLeadingCharacters()
        {
            var options = new PositionRemovalOptions
            {
                Mode = PositionRemovalMode.FromLeft,
                Count = 4
            };
            string result = PositionRemover.Remove("PREFIX_Data_2026", options);
            Assert.AreEqual("IX_Data_2026", result);
        }

        [TestMethod]
        public void PositionRemover_FromRight_RemovesTrailingCharacters()
        {
            var options = new PositionRemovalOptions
            {
                Mode = PositionRemovalMode.FromRight,
                Count = 5
            };
            string result = PositionRemover.Remove("Filename.xlsx", options);
            Assert.AreEqual("Filename", result);
        }

        [TestMethod]
        public void PositionRemover_FromMiddle_RemovesSubrange()
        {
            var options = new PositionRemovalOptions
            {
                Mode = PositionRemovalMode.FromMiddle,
                StartIndex = 4, // 1-based: starts at 4th char ('D')
                Count = 3       // removes 'DEF'
            };
            string result = PositionRemover.Remove("ABCDEFGHI", options);
            Assert.AreEqual("ABCGHI", result);
        }

        [TestMethod]
        public void PositionRemover_BeforeDelimiter_RemovesPrefix()
        {
            var options = new PositionRemovalOptions
            {
                Mode = PositionRemovalMode.BeforeDelimiter,
                Delimiter = ":"
            };
            string result = PositionRemover.Remove("KEY_ID:TARGET_VALUE", options);
            Assert.AreEqual("TARGET_VALUE", result);
        }

        [TestMethod]
        public void PositionRemover_AfterDelimiter_RemovesSuffix()
        {
            var options = new PositionRemovalOptions
            {
                Mode = PositionRemovalMode.AfterDelimiter,
                Delimiter = ":"
            };
            string result = PositionRemover.Remove("KEY_ID:TARGET_VALUE", options);
            Assert.AreEqual("KEY_ID", result);
        }

        [TestMethod]
        public void PositionRemover_AfterDelimiter_RemovesDateSuffix()
        {
            var options = new PositionRemovalOptions
            {
                Mode = PositionRemovalMode.AfterDelimiter,
                Delimiter = "-"
            };
            string result = PositionRemover.Remove("2026-10-20 14:02:38.019", options);
            Assert.AreEqual("2026", result);
        }

        [TestMethod]
        public void PositionRemover_BeforeDelimiter_RemovesDatePrefix()
        {
            var options = new PositionRemovalOptions
            {
                Mode = PositionRemovalMode.BeforeDelimiter,
                Delimiter = "-"
            };
            string result = PositionRemover.Remove("2026-10-20 14:02:38.019", options);
            Assert.AreEqual("10-20 14:02:38.019", result);
        }

        [TestMethod]
        public void PositionRemover_MultiAreaRanges_ProcessedIndependently()
        {
            var area1 = new object?[,]
            {
                { "2026-10-31 0" },
                { "2026-10-31 0" }
            };

            var area2 = new object?[,]
            {
                { "2026-10-26 10:00:56.56" },
                { "2026-10-26 09:50:42.486" }
            };

            var options = new PositionRemovalOptions
            {
                Mode = PositionRemovalMode.FromRight,
                Count = 2
            };

            var res1 = PositionRemover.RemoveRange(area1, options);
            var res2 = PositionRemover.RemoveRange(area2, options);

            Assert.AreEqual("2026-10-31", res1[0, 0]);
            Assert.AreEqual("2026-10-31", res1[1, 0]);
            Assert.AreEqual("2026-10-26 10:00:56.", res2[0, 0]);
            Assert.AreEqual("2026-10-26 09:50:42.4", res2[1, 0]);
        }

        #endregion
    }
}
