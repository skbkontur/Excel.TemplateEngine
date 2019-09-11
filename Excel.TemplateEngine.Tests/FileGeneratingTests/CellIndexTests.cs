using Excel.TemplateEngine.FileGenerating.Implementation;

using NUnit.Framework;

namespace Excel.TemplateEngine.Tests.FileGeneratingTests
{
    [TestFixture]
    public class CellIndexTests
    {
        [Test]
        public void CellIndexConstructorsTest()
        {
            var a = new ExcelCellIndex("A1");
            Assert.AreEqual(a.RowIndex, 1);
            Assert.AreEqual(a.ColumnIndex, 1);

            var b = new ExcelCellIndex(1, 1);
            Assert.AreEqual(b.CellReference, "A1");

            var c = new ExcelCellIndex("AAA23");
            Assert.AreEqual(c.RowIndex, 23);
            Assert.AreEqual(c.ColumnIndex, 703);
        }

        [Test]
        public void CellIndexSubtractionTest()
        {
            var a = new ExcelCellIndex("A1");
            var b = new ExcelCellIndex("A1");

            Assert.AreEqual("A1", a.Subtract(b).CellReference);

            b = new ExcelCellIndex("ABC345");
            Assert.AreEqual("ABC345", b.Subtract(a).CellReference);
        }

        [Test]
        public void CellIndexAdditionTest()
        {
            var a = new ExcelCellIndex("A1");
            var b = new ExcelCellIndex("A1");

            Assert.AreEqual("A1", a.Add(b).CellReference);
        }
    }
}