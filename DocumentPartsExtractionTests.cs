using System.IO;
using System.Linq;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;

namespace SKBKontur.Catalogue.Core.Tests.ExcelFileGeneratorTests
{
    [TestFixture]
    public class DocumentPartsExtractionTests
    {
        [Test]
        public void CellsInRangeTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(fileName));
            var worksheet = document.GetWorksheet(0);
            var cells = worksheet.GetSortedCellsInRange("A34", "B38").ToArray();

            Assert.AreEqual(10, cells.Count());
            Assert.AreEqual("Value:String:Name", cells[1].GetStringValue());
            Assert.AreEqual("Value:String:Address", cells[9].GetStringValue());

            document.Dispose();
        }

        private const string fileName = @"ExcelFileGeneratorTests\Files\template.xlsx";
    }
}