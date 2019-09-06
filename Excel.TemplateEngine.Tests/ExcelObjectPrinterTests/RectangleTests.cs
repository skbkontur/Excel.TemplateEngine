using NUnit.Framework;

using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class RectangleTests
    {
        [Test]
        public void ByTwoPositionsConstructorTest()
        {
            var rect = new Rectangle(new CellPosition(2, 1), new CellPosition(3, 5));

            Assert.AreEqual(5, rect.Size.Width);
            Assert.AreEqual(2, rect.Size.Height);
        }

        [Test]
        public void ByPositionAndSizeConstructorTest()
        {
            var rect = new Rectangle(new CellPosition(2, 1), new ObjectSize(5, 2));

            Assert.AreEqual(3, rect.LowerRight.RowIndex);
            Assert.AreEqual(5, rect.LowerRight.ColumnIndex);
        }

        [Test]
        public void IntersectionTest1()
        {
            var rect1 = new Rectangle(new CellPosition(1, 2), new ObjectSize(4, 3));
            var rect2 = new Rectangle(new CellPosition(2, 3), new ObjectSize(2, 4));

            Assert.AreEqual(true, rect1.Intersects(rect2));
        }

        [Test]
        public void IntersectionTest2()
        {
            var rect1 = new Rectangle(new CellPosition(1, 2), new ObjectSize(4, 3));
            var rect2 = new Rectangle(new CellPosition(0, 3), new ObjectSize(4, 10));

            Assert.AreEqual(true, rect1.Intersects(rect2));
        }

        [Test]
        public void IntersectionTest3()
        {
            var rect1 = new Rectangle(new CellPosition(1, 2), new ObjectSize(2, 1));
            var rect2 = new Rectangle(new CellPosition(2, 3), new ObjectSize(2, 4));

            Assert.AreEqual(false, rect1.Intersects(rect2));
        }

        [Test]
        public void IntersectionTest4()
        {
            var rect1 = new Rectangle(new CellPosition(1, 2), new ObjectSize(2, 1));
            var rect2 = new Rectangle(new CellPosition(1, 3), new ObjectSize(2, 1));

            Assert.AreEqual(true, rect1.Intersects(rect2));
        }

        [Test]
        public void IntersectionTest5()
        {
            var rect1 = new Rectangle(new CellPosition(1, 2), new ObjectSize(1, 1));
            var rect2 = new Rectangle(new CellPosition(2, 3), new ObjectSize(1, 1));

            Assert.AreEqual(false, rect1.Intersects(rect2));
        }

        [Test]
        public void IntersectionTest6()
        {
            var rect1 = new Rectangle(new CellPosition(1, 1), new CellPosition(4, 4));
            var rect2 = new Rectangle(new CellPosition(2, 2), new CellPosition(3, 3));

            Assert.AreEqual(true, rect1.Intersects(rect2));
        }

        [Test]
        public void IntersectionTest7()
        {
            var rect1 = new Rectangle(new CellPosition(1, 1), new CellPosition(4, 4));
            var rect2 = new Rectangle(new CellPosition(1, 1), new CellPosition(1, 2));

            Assert.AreEqual(true, rect1.Intersects(rect2));
        }

        [Test]
        public void IntersectionTest8()
        {
            var rect1 = new Rectangle(new CellPosition(1, 1), new CellPosition(4, 4));
            var rect2 = new Rectangle(new CellPosition(1, 1), new CellPosition(2, 1));

            Assert.AreEqual(true, rect1.Intersects(rect2));
        }

        [Test]
        public void IntersectionTest9()
        {
            var rect1 = new Rectangle(new CellPosition(1, 1), new CellPosition(4, 4));
            var rect2 = new Rectangle(new CellPosition(4, 3), new CellPosition(4, 4));

            Assert.AreEqual(true, rect1.Intersects(rect2));
        }

        [Test]
        public void IntersectionTest10()
        {
            var rect1 = new Rectangle(new CellPosition(1, 1), new CellPosition(4, 4));
            var rect2 = new Rectangle(new CellPosition(3, 4), new CellPosition(4, 4));

            Assert.AreEqual(true, rect1.Intersects(rect2));
        }

        [Test]
        public void PositionContainingTest()
        {
            var rect = new Rectangle(new CellPosition(1, 1), new CellPosition(45, 36));
            var position = new CellPosition(1, 3);

            Assert.AreEqual(true, rect.Contains(position));
        }
    }
}