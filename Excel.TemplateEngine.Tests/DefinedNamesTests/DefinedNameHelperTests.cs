using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;

using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

namespace SkbKontur.Excel.TemplateEngine.Tests.DefinedNamesTests
{
    public class DefinedNameHelperTests
    {
        [TestCaseSource(nameof(GetIsBelongsToSheetTestData))]
        public void IsBelongsToSheetTest_ShouldBelong(string definedNameInnerText, string sheetName)
        {
            var definedName = new DefinedName(definedNameInnerText);
            var sheet = new Sheet {Name = sheetName};

            definedName.BelongsToSheet(sheet).Should().BeTrue();
        }

        [TestCase("Лист1!$A$1", "Лист 1")]
        [TestCase("'Лист 1'!$A$1", "Лист1")]
        public void IsBelongsToSheetTest_ShouldNotBelong(string definedNameInnerText, string sheetName)
        {
            var definedName = new DefinedName(definedNameInnerText);
            var sheet = new Sheet {Name = sheetName};

            definedName.BelongsToSheet(sheet).Should().BeFalse();
        }

        [TestCaseSource(nameof(GetReplaceSheetNameInFormulaTestData))]
        public void ReplaceSheetNameInFormulaTest(string innerText, string name, string expectedDefinedNameInnerText)
        {
            var definedName = new DefinedName(innerText);

            var updatedDefinedName = definedName.ReplaceSheetNameInFormula(name);
            updatedDefinedName.InnerText.Should().Be(expectedDefinedNameInnerText);
        }

        private static IEnumerable<TestCaseData> GetIsBelongsToSheetTestData()
        {
            return new[]
                {
                    new TestCaseData("Лист1!$A$1", "Лист1"),
                    new TestCaseData("Лист_1!$A$1", "Лист_1"),
                    new TestCaseData("'Лист''1'!$A$1", "Лист'1"),
                    new TestCaseData("'Лист 1'!$A$1", "Лист 1"),
                    new TestCaseData("'Лист1!'!$A$1", "Лист1!"),
                    new TestCaseData("'Лист1 !'!$A$1", "Лист1 !"),
                    new TestCaseData("'Лист1 ''!'!$A$1", "Лист1 '!"),
                    new TestCaseData("' Лист1 '!$A$1", " Лист1 "),
                    new TestCaseData("Лист1!$F$18:$F$201", "Лист1"),
                    new TestCaseData("' Лист1 ''!'!$F$18:$F$201", " Лист1 '!")
                };
        }

        private static IEnumerable<TestCaseData> GetReplaceSheetNameInFormulaTestData()
        {
            return new[]
                {
                    new TestCaseData("Лист1!$A$1", "Лист2", "Лист2!$A$1"),

                    // Пробел
                    new TestCaseData("Лист1!$A$1", "Лист 1", "'Лист 1'!$A$1"),

                    // Спецсимвол
                    new TestCaseData("Лист1!$A$1", "Лист_1", "Лист_1!$A$1"),
                    new TestCaseData("Лист1!$A$1", "_Лист1", "_Лист1!$A$1"),
                    new TestCaseData("Лист1!$A$1", "Лист1!", "'Лист1!'!$A$1"),
                    new TestCaseData("Лист1!$A$1", "Ли%ст1", "'Ли%ст1'!$A$1"),
                    new TestCaseData("Лист1!$A$1", "$Лист1$", "'$Лист1$'!$A$1"),

                    // Первый символ - цифра
                    new TestCaseData("Лист1!$A$1", "1Лист", "'1Лист'!$A$1"),
                    new TestCaseData("Лист1!$A$1", "034Лист", "'034Лист'!$A$1"),

                    // Одинарная кавычка
                    new TestCaseData("Лист1!$A$1", "Лист'1", "'Лист''1'!$A$1"),

                    // A1-нотация
                    new TestCaseData("Лист1!$A$1", "A1", "'A1'!$A$1"),
                    new TestCaseData("Лист1!$A$1", "a1", "'a1'!$A$1"),
                    new TestCaseData("Лист1!$A$1", "ACB12", "'ACB12'!$A$1"),
                    new TestCaseData("Лист1!$A$1", "ACb12", "'ACb12'!$A$1")
                };
        }
    }
}