using System.IO;

using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.FileGenerating;
using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator;

using Vostok.Logging.Console;

namespace SkbKontur.Excel.TemplateEngine.Tests.FileGeneratingTests
{
    public class ExcelStyleTests : FileBasedTestBase
    {
        [Test]
        public void CellColorExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            style.FillStyle.Color.Red.Should().Be(128);
            style.FillStyle.Color.Green.Should().Be(129);
            style.FillStyle.Color.Blue.Should().Be(200);
        }

        [Test]
        public void CellFontExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            style.FontStyle.Bold.Should().BeTrue();
            style.FontStyle.Underlined.Should().BeTrue();
            style.FontStyle.Size.Should().Be(20);

            style.FontStyle.Color.Red.Should().Be(122);
            style.FontStyle.Color.Green.Should().Be(200);
            style.FontStyle.Color.Blue.Should().Be(20);
        }

        [Test]
        public void CellNumberFormatExtractionTest_Custom()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            style.NumberingFormat.Code.Should().Be("0.0000");
        }

        /// <summary>
        ///     Test use one of standard numbering formats being stored without format code.
        ///     To see all supported standard formats look for ExcelDocumentStyle.standardNumberingFormatsId
        /// </summary>
        [Test]
        public void CellNumberFormatExtractionTest_Standard()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A4"));
            var style = cell.GetStyle();

            style.NumberingFormat.Id.Should().Be(49);
            style.NumberingFormat.Code.Should().BeNull();
        }

        [Test]
        public void CellNumberFormatCopyAfterRenderTest()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsx")), logger))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("A1"), logger), new Style(template.GetCell(new CellPosition("A1"))));

                // Model is required for printing cells defined by Template:RootTemplate
                var dummyModel = new {A = 1, B = "asdf"};
                templateEngine.Render(tableBuilder, dummyModel);

                var worksheet = targetDocument.GetWorksheet(0);
                var customStyle = worksheet.GetCell(new ExcelCellIndex("A1")).GetStyle();
                customStyle.NumberingFormat.Code.Should().Be("0.0000");

                var standardStyle = worksheet.GetCell(new ExcelCellIndex("A4")).GetStyle();
                standardStyle.NumberingFormat.Id.Should().Be(49);
                standardStyle.NumberingFormat.Code.Should().BeNull();
            }
        }

        [Test]
        public void CellBorderFormatExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            style.BordersStyle.BottomBorder.BorderType.Should().Be(ExcelBorderType.Double);
            style.BordersStyle.BottomBorder.Color.Red.Should().Be(255);
            style.BordersStyle.BottomBorder.Color.Green.Should().Be(0);
            style.BordersStyle.BottomBorder.Color.Blue.Should().Be(255);
        }

        [TestCase(ExcelHorizontalAlignment.Left, "A1")]
        [TestCase(ExcelHorizontalAlignment.Fill, "B1")]
        public void CellAlignmentExtractionTest(ExcelHorizontalAlignment horizontalAlignment, string cellIndex)
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex(cellIndex));
            var style = cell.GetStyle();

            style.Alignment.WrapText.Should().BeTrue();
            style.Alignment.HorizontalAlignment.Should().Be(horizontalAlignment);
        }

        [Test]
        public void NullCellStyleAccessTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A2"));
            cell.GetStyle();
        }

        private readonly ConsoleLog logger = new ConsoleLog();
    }
}