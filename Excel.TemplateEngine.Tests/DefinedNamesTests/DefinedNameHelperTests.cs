using DocumentFormat.OpenXml.Spreadsheet;

using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.FileGenerating.Helpers;

namespace SkbKontur.Excel.TemplateEngine.Tests.DefinedNamesTests
{
    public class DefinedNameHelperTests
    {
        [TestCase("Sheet1!$A$1", "Sheet1")]
        [TestCase("S!h$ee%t1!$A$1", "S!h$ee%t1")]
        public void IsBelongsToSheetTest_ShouldBelong(string definedNameInnerText, string sheetName)
        {
            var definedName = new DefinedName(definedNameInnerText);
            var sheet = new Sheet {Name = sheetName};

            definedName.IsBelongsToSheet(sheet).Should().BeTrue();
        }

        [TestCase("Sheet1!$A$1", "Sheet2")]
        [TestCase("S!h$ee%t1!$A$1", "S!h$ee%t2")]
        public void IsBelongsToSheetTest_ShouldNotBelong(string definedNameInnerText, string sheetName)
        {
            var definedName = new DefinedName(definedNameInnerText);
            var sheet = new Sheet {Name = sheetName};

            definedName.IsBelongsToSheet(sheet).Should().BeFalse();
        }

        [TestCase("Sheet1!$A$1", "Sheet2", "Sheet2!$A$1")]
        [TestCase("She!!!et1!$A$1", "Sheet2", "Sheet2!$A$1")]
        public void UpdateLinkSheetName(string definedNameInnerText, string sheetName, string expectedDefinedNameInnerText)
        {
            var definedName = new DefinedName(definedNameInnerText);

            var updatedDefinedName = definedName.UpdateLinkSheetName(sheetName);
            updatedDefinedName.InnerText.Should().Be(expectedDefinedNameInnerText);
        }
    }
}