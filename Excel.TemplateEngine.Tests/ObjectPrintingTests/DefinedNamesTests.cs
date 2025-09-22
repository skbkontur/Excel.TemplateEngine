using System.IO;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml.Spreadsheet;

using NUnit.Framework;

using FluentAssertions;

using SkbKontur.Excel.TemplateEngine.FileGenerating;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator;

using Vostok.Logging.Console;

namespace SkbKontur.Excel.TemplateEngine.Tests.ObjectPrintingTests
{
    public class DefinedNamesTests : FileBasedTestBase
    {
        [Test]
        public void TestCopyDefinedNames()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("customCellNames.xlsx")), logger))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsm")), logger))
            {
                targetDocument.CopyDefinedNamesFrom(templateDocument);

                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"), logger);
                var tableBuilder = new TableBuilder(target, tableNavigator, new Style(template.GetCell(new CellPosition("A1"))));
                templateEngine.Render(tableBuilder, new {Test = "b"});

                var expectedData = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("customCellNamesResult.xlsm")), logger);
                var actualData = ExcelDocumentFactory.CreateFromTemplate(targetDocument.CloseAndGetDocumentBytes(), logger);

                var expectedDefinedNames = expectedData.GetDefinedNames().InnerXml;
                var actualDefinedNames = actualData.GetDefinedNames().InnerXml;

                expectedDefinedNames.Should().Match(actualDefinedNames);

                var expectedString = expectedData.ToString();
                var actualString = actualData.ToString();

                var expectedBytes = Encoding.Default.GetBytes(expectedString);
                var actualBytes = Encoding.Default.GetBytes(actualString);

                actualBytes.Should().Equal(expectedBytes);
            }
        }

        /// <summary>
        ///     Проверка элементов DefinedNames при переименовании листа
        ///     1. Без обновления DefinedNames, объекты DefinedNames остаются в документе, но содержат в своей ссылке старое имя листа
        ///     2. С обновлением DefinedNames, объекты DefinedNames остаются в документе и содержат в своей ссылке новое имя листа
        /// </summary>
        [TestCase(false)]
        [TestCase(true)]
        public void TestDefinedNamesUpdatedAfterSheetRename(bool updateDefinedNames)
        {
            using (var templateDocument = ExcelDocumentFactory
                       .CreateFromTemplate(File.ReadAllBytes(GetFilePath("definedNamesUpdatedAfterSheetRename.xlsx")), logger))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsm")), logger))
            {
                targetDocument.CopyDefinedNamesFrom(templateDocument);

                var templateDefinedNames = templateDocument.GetDefinedNames();
                var targetDefinedNames = targetDocument.GetDefinedNames();

                targetDefinedNames.Count().Should().Be(templateDefinedNames.Count());
                targetDefinedNames.InnerText.Should().Match(templateDefinedNames.InnerText);

                var oldSheetName = targetDocument.GetWorksheetName(0);
                var newSheetName = "NewName";
                targetDocument.RenameWorksheet(0, newSheetName, updateDefinedNames);

                var targetDocumentAfterRenamingSheet = ExcelDocumentFactory
                    .CreateFromTemplate(targetDocument.CloseAndGetDocumentBytes(), logger);
                targetDefinedNames = targetDocumentAfterRenamingSheet.GetDefinedNames();

                targetDefinedNames.Should().NotBeNull();
                targetDefinedNames.Count().Should().Be(templateDefinedNames.Count());
                foreach (var definedName in targetDefinedNames.Elements<DefinedName>())
                {
                    var sheetName = GetDefinedNameAddress(definedName);
                    sheetName.sheetName.Should().Be(updateDefinedNames ? newSheetName : oldSheetName);
                }
            }
        }

        private static (string sheetName, string cellAddress) GetDefinedNameAddress(DefinedName definedName)
        {
            var exclamationIndex = definedName.InnerText.LastIndexOf('!');
            var sheetName = definedName.InnerText.Substring(0, exclamationIndex);
            var cellAddress = definedName.InnerText.Substring(exclamationIndex + 1);

            return (sheetName, cellAddress);
        }

        private readonly ConsoleLog logger = new ConsoleLog();
    }
}