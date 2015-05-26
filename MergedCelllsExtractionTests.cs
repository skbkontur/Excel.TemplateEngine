using System.IO;
using System.Linq;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation;

namespace SKBKontur.Catalogue.Core.Tests.ExcelFileGeneratorTests
{
    [TestFixture]
    public class MergedCelllsExtractionTests
    {
        [Test]
        public void MergedCellsExtractionTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(templateFileName));
            var worksheet = document.GetWorksheet(0);

            var cells = worksheet.MergedCells.ToArray();

            Assert.AreEqual(2, cells.Length);
            Assert.AreEqual("B32", cells[0].Item1.CellReference);
            Assert.AreEqual("C35", cells[0].Item2.CellReference);
            Assert.AreEqual("I2", cells[1].Item1.CellReference);
            Assert.AreEqual("J3", cells[1].Item2.CellReference);

            document.Dispose();
        }

        [Test]
        public void MergedCellsExtractionEmptyTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(emptyFileName));
            var worksheet = document.GetWorksheet(0);

            var cells = worksheet.MergedCells.ToArray();

            Assert.AreEqual(0, cells.Length);

            document.Dispose();
        }

        private const string templateFileName = @"ExcelFileGeneratorTests\Files\template.xlsx";
        private const string emptyFileName = @"ExcelFileGeneratorTests\Files\empty.xlsx";
    }
}