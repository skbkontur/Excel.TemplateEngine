using System.IO;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;
using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation;

namespace SKBKontur.Catalogue.Core.Tests.ExcelFileGeneratorTests
{
    [TestFixture]
    public class ExcelStyleTests
    {
        [Test]
        public void CellColorExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(templateFileName));
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            Assert.AreEqual(128, style.FillStyle.Color.Red);
            Assert.AreEqual(129, style.FillStyle.Color.Green);
            Assert.AreEqual(200, style.FillStyle.Color.Blue);
        }

        [Test]
        public void CellFontExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(templateFileName));
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
            var templateDoucument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(templateFileName));
            var cell = templateDoucument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            Assert.AreEqual(4, style.NumberingFormat.Precision);
        }

        [Test]
        public void CellBorderFormatExtractionTest()
        {
            var templateDoucument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(templateFileName));
            var cell = templateDoucument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            Assert.AreEqual(ExcelBorderType.Double, style.BordersStyle.BottomBorder.BorderType);
            Assert.AreEqual(255, style.BordersStyle.BottomBorder.Color.Red);
            Assert.AreEqual(0, style.BordersStyle.BottomBorder.Color.Green);
            Assert.AreEqual(255, style.BordersStyle.BottomBorder.Color.Blue);
        }

        [Test]
        public void CellAlignmentExtractionTest()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(templateFileName));
            var cell = templateDocument.GetWorksheet(0).GetCell(new ExcelCellIndex("A1"));
            var style = cell.GetStyle();

            Assert.IsTrue(style.Alignment.WrapText);
            Assert.AreEqual(ExcelHorizontalAlignment.Left, style.Alignment.HorizontalAlignment);
        }

        [Test]
        public void NullCellStyleAccessTest()
        {
            var templateDoucument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(templateFileName));
            var cell = templateDoucument.GetWorksheet(0).GetCell(new ExcelCellIndex("A2"));
            cell.GetStyle();
        }

        private const string templateFileName = @"ExcelFileGeneratorTests\Files\template.xlsx";
    }
}