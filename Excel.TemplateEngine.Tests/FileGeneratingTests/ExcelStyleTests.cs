using System.IO;

using Excel.TemplateEngine.FileGenerating;
using Excel.TemplateEngine.FileGenerating.DataTypes;
using Excel.TemplateEngine.FileGenerating.Implementation;

using NUnit.Framework;

namespace Excel.TemplateEngine.Tests.FileGeneratingTests
{
    public class ExcelStyleTests : FileBasedTestBase
    {
        [Test]
        public void CellColorExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), Log.DefaultLogger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            Assert.AreEqual(128, style.FillStyle.Color.Red);
            Assert.AreEqual(129, style.FillStyle.Color.Green);
            Assert.AreEqual(200, style.FillStyle.Color.Blue);
        }

        [Test]
        public void CellFontExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), Log.DefaultLogger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            Assert.IsTrue(style.FontStyle.Bold);
            Assert.IsTrue(style.FontStyle.Underlined);
            Assert.AreEqual(20, style.FontStyle.Size);

            Assert.AreEqual(122, style.FontStyle.Color.Red);
            Assert.AreEqual(200, style.FontStyle.Color.Green);
            Assert.AreEqual(20, style.FontStyle.Color.Blue);
        }

        [Test]
        public void CellNumberFormatExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), Log.DefaultLogger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            Assert.AreEqual("0.0000", style.NumberingFormat.FormatCode);
        }

        [Test]
        public void CellBorderFormatExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), Log.DefaultLogger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            Assert.AreEqual(ExcelBorderType.Double, style.BordersStyle.BottomBorder.BorderType);
            Assert.AreEqual(255, style.BordersStyle.BottomBorder.Color.Red);
            Assert.AreEqual(0, style.BordersStyle.BottomBorder.Color.Green);
            Assert.AreEqual(255, style.BordersStyle.BottomBorder.Color.Blue);
        }

        [Test]
        public void CellAlignmentExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), Log.DefaultLogger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            Assert.IsTrue(style.Alignment.WrapText);
            Assert.AreEqual(ExcelHorizontalAlignment.Left, style.Alignment.HorizontalAlignment);
        }

        [Test]
        public void NullCellStyleAccessTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), Log.DefaultLogger);
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A2"));
            cell.GetStyle();
        }
    }
}