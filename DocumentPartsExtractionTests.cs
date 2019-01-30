using System.IO;
using System.Linq;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation;

namespace SKBKontur.Catalogue.Core.Tests.ExcelFileGeneratorTests
{
    public class DocumentPartsExtractionTests : FileBasedTestBase
    {
        [Test]
        public void CellsInRangeTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var worksheet = document.GetWorksheet(0);

            var upperLeft = new ExcelCellIndex("A22");
            var lowerRight = new ExcelCellIndex("A24");
            var cells = worksheet.GetSortedCellsInRange(upperLeft, lowerRight).ToArray();

            Assert.AreEqual(3, cells.Count());
            Assert.AreEqual("Value:String:Name", cells[1].GetStringValue());
            Assert.AreEqual("Value:String:Address", cells[2].GetStringValue());

            document.Dispose();
        }

        [Test]
        public void GetCellTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var worksheet = document.GetWorksheet(0);

            var position = new ExcelCellIndex("B9");
            var cell = worksheet.GetCell(position);
            Assert.AreEqual("Model:Document:B10:C11", cell.GetStringValue());

            document.Dispose();
        }

        [Test]
        public void SearchCellsByTextTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger);
            var worksheet = document.GetWorksheet(0);

            var cell = worksheet.SearchCellsByText("Value:String").FirstOrDefault();
            Assert.AreNotEqual(null, cell);
// ReSharper disable PossibleNullReferenceException
            Assert.AreEqual("Value:String:Name", cell.GetStringValue());
// ReSharper restore PossibleNullReferenceException

            document.Dispose();
        }
    }
}