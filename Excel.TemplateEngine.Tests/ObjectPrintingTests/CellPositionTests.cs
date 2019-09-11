using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

using NUnit.Framework;

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

            Assert.AreEqual(c.Width, 0);
            Assert.AreEqual(c.Height, 0);
        }

        [Test]
        public void CellIndexSubtractionTest2()
        {
            var a = new CellPosition(344, 856);
            var b = new CellPosition(123, 345);
            var c = a.Subtract(b);

            Assert.AreEqual(c.Width, 511);
            Assert.AreEqual(c.Height, 221);
        }

        [Test]
        public void ObjectSizeAdditionTest()
        {
            var a = new ObjectSize(23, 22);
            var b = new ObjectSize(12, 15);
            var c = a.Add(b);

            Assert.AreEqual(c.Width, 35);
            Assert.AreEqual(c.Height, 37);
        }

        [Test]
        public void ObjectSizeSubtractionTest()
        {
            var a = new ObjectSize(1, 22);
            var b = new ObjectSize(1, 11);
            var c = a.Subtract(b);

            Assert.AreEqual(c.Width, 0);
            Assert.AreEqual(c.Height, 11);
        }

        [Test]
        public void CellIndexAndObjectSizeAdditionTest1()
        {
            var a = new CellPosition("A1");
            var b = new ObjectSize(0, 0);
            var c = a.Add(b);

            Assert.AreEqual(c.RowIndex, 1);
            Assert.AreEqual(c.ColumnIndex, 1);
        }

        [Test]
        public void CellIndexAndObjectSizeAdditionTest2()
        {
            var a = new CellPosition("A1");
            var b = new ObjectSize(2, 2);
            var c = a.Add(b);

            Assert.AreEqual(c.RowIndex, 3);
            Assert.AreEqual(c.ColumnIndex, 3);
        }

        [Test]
        public void ToRelativeCoordinatesConversionTest()
        {
            var pos = new CellPosition(3, 6);
            var origin = new CellPosition(2, 2);

            var relativePos = pos.ToRelativeCoordinates(origin);

            Assert.AreEqual(relativePos.RowIndex, 2);
            Assert.AreEqual(relativePos.ColumnIndex, 5);
        }

        [Test]
        public void ToGlobalCoordinatesConversionTest()
        {
            var pos = new CellPosition(2, 5);
            var origin = new CellPosition(2, 2);

            var globalPos = pos.ToGlobalCoordinates(origin);

            Assert.AreEqual(globalPos.RowIndex, 3);
            Assert.AreEqual(globalPos.ColumnIndex, 6);
        }
    }
}