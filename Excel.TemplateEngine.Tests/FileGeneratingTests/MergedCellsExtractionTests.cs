using System.IO;

using NUnit.Framework;

namespace Excel.TemplateEngine.Tests.FileGeneratingTests
{
    [TestFixture]
    public class MergedCellsExtractionTests : FileBasedTestBase
    {
        [Test]
        public void MergedCellsExtractionTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), Log.DefaultLogger);
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
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsx")), Log.DefaultLogger);
            var worksheet = document.GetWorksheet(0);

            var cells = worksheet.MergedCells.ToArray();

            Assert.AreEqual(0, cells.Length);

            document.Dispose();
        }
    }
}