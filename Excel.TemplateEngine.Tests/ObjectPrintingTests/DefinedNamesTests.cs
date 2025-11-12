using System.IO;
using System.Text;

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
        ///     Проверка корректности элементов DefinedNames после переименования листа
        /// </summary>
        [Test]
        public void TestDefinedNamesUpdatedAfterSheetRename()
        {
            using (var document = ExcelDocumentFactory
                       .CreateFromTemplate(File.ReadAllBytes(GetFilePath("definedNamesUpdatedAfterSheetRename.xlsx")),
                                           logger))
            {
                document.RenameWorksheet(0, "НазваниеБезПробелов", true);
                document.RenameWorksheet(1, "НазваниеС Пробелом", true);
                document.RenameWorksheet(2, "НазваниеСо_Спецсимволом", true);
                document.RenameWorksheet(3, "НазваниеСо+Спецсимволом", true);
                document.RenameWorksheet(4, "1НазваниеСЦифройВначале", true);
                document.RenameWorksheet(5, "НазваниеС'ОдинарнойКовычкой", true);
                document.RenameWorksheet(6, "AN1", true);

                var expectedData = ExcelDocumentFactory.CreateFromTemplate(
                    File.ReadAllBytes(GetFilePath("definedNamesUpdatedAfterSheetRenameResult.xlsx")),
                    logger);
                var actualData = ExcelDocumentFactory.CreateFromTemplate(
                    document.CloseAndGetDocumentBytes(),
                    logger);

                var expectedDefinedNames = expectedData.GetDefinedNames().InnerXml;
                var actualDefinedNames = actualData.GetDefinedNames().InnerXml;

                expectedDefinedNames.Should().Match(actualDefinedNames);
            }
        }

        private readonly ConsoleLog logger = new ConsoleLog();
    }
}