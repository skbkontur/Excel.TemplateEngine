using Excel.TemplateEngine.FileGenerating.DataTypes;

using FluentAssertions;

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
            a.RowIndex.Should().Be(1);
            a.ColumnIndex.Should().Be(1);

            var b = new ExcelCellIndex(1, 1);
            b.CellReference.Should().Be("A1");

            var c = new ExcelCellIndex("AAA23");
            c.RowIndex.Should().Be(23);
            c.ColumnIndex.Should().Be(703);
        }

        [Test]
        public void CellIndexSubtractionTest()
        {
            var a = new ExcelCellIndex("A1");
            var b = new ExcelCellIndex("A1");

            "A1".Should().Be(a.Subtract(b).CellReference);

            b = new ExcelCellIndex("ABC345");
            "ABC345".Should().Be(b.Subtract(a).CellReference);
        }

        [Test]
        public void CellIndexAdditionTest()
        {
            var a = new ExcelCellIndex("A1");
            var b = new ExcelCellIndex("A1");

            a.Add(b).CellReference.Should().Be("A1");
        }
    }
}