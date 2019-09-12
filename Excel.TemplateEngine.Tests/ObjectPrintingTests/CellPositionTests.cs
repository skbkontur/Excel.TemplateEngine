using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests
{
    public class CellPositionTests
    {
        [Test]
        public void CellIndexSubtractionTest1()
        {
            var a = new CellPosition(1, 1);
            var b = new CellPosition("A1");
            var c = a.Subtract(b);

            c.Width.Should().Be(0);
            c.Height.Should().Be(0);
        }

        [Test]
        public void CellIndexSubtractionTest2()
        {
            var a = new CellPosition(344, 856);
            var b = new CellPosition(123, 345);
            var c = a.Subtract(b);

            c.Width.Should().Be(511);
            c.Height.Should().Be(221);
        }

        [Test]
        public void ObjectSizeAdditionTest()
        {
            var a = new ObjectSize(23, 22);
            var b = new ObjectSize(12, 15);
            var c = a.Add(b);

            c.Width.Should().Be(35);
            c.Height.Should().Be(37);
        }

        [Test]
        public void ObjectSizeSubtractionTest()
        {
            var a = new ObjectSize(1, 22);
            var b = new ObjectSize(1, 11);
            var c = a.Subtract(b);

            c.Width.Should().Be(0);
            c.Height.Should().Be(11);
        }

        [Test]
        public void CellIndexAndObjectSizeAdditionTest1()
        {
            var a = new CellPosition("A1");
            var b = new ObjectSize(0, 0);
            var c = a.Add(b);

            c.RowIndex.Should().Be(1);
            c.ColumnIndex.Should().Be(1);
        }

        [Test]
        public void CellIndexAndObjectSizeAdditionTest2()
        {
            var a = new CellPosition("A1");
            var b = new ObjectSize(2, 2);
            var c = a.Add(b);

            c.RowIndex.Should().Be(3);
            c.ColumnIndex.Should().Be(3);
        }

        [Test]
        public void ToRelativeCoordinatesConversionTest()
        {
            var pos = new CellPosition(3, 6);
            var origin = new CellPosition(2, 2);

            var relativePos = pos.ToRelativeCoordinates(origin);

            relativePos.RowIndex.Should().Be(2);
            relativePos.ColumnIndex.Should().Be(5);
        }

        [Test]
        public void ToGlobalCoordinatesConversionTest()
        {
            var pos = new CellPosition(2, 5);
            var origin = new CellPosition(2, 2);

            var globalPos = pos.ToGlobalCoordinates(origin);

            globalPos.RowIndex.Should().Be(3);
            globalPos.ColumnIndex.Should().Be(6);
        }
    }
}