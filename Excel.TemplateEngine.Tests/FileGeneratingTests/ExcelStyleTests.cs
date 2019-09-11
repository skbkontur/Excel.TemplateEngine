using System.IO;

using Excel.TemplateEngine.FileGenerating;
using Excel.TemplateEngine.FileGenerating.DataTypes;
using Excel.TemplateEngine.FileGenerating.Implementation;

using FluentAssertions;

using NUnit.Framework;

using Vostok.Logging.Console;

namespace Excel.TemplateEngine.Tests.FileGeneratingTests
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
        public void CellNumberFormatExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            style.NumberingFormat.FormatCode.Should().Be("0.0000");
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

        [Test]
        public void CellAlignmentExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            style.Alignment.WrapText.Should().BeTrue();
            style.Alignment.HorizontalAlignment.Should().Be(ExcelHorizontalAlignment.Left);
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