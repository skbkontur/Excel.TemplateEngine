using NUnit.Framework;

using SKBKontur.Catalogue.ExcelObjectPrinter.FakeDocumentPrimitivesImplementation;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class TableBuilderTests
    {
        [Test]
        public void TableBuilderAtomicValuePrintingTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var tableBuilder = new TableBuilder(table, new CellPosition("A1"));

            tableBuilder.RenderAtomicValue("Test");
            tableBuilder.MoveToNextColumn();
            Assert.AreEqual("Test", table.GetCell(new CellPosition("A1")).StringValue);
            Assert.AreEqual("B1", tableBuilder.CurrentState.Cursor.CellReference);
            Assert.AreEqual(0, tableBuilder.CurrentState.GlobalHeight);
            Assert.AreEqual(1, tableBuilder.CurrentState.GlobalWidth);
            Assert.AreEqual(0, tableBuilder.CurrentState.CurrentLayerHeight);
            Assert.AreEqual(1, tableBuilder.CurrentState.CurrentLayerStartRowIndex);

            tableBuilder.RenderAtomicValue("tseT");
            tableBuilder.MoveToNextColumn();
            Assert.AreEqual("tseT", table.GetCell(new CellPosition("B1")).StringValue);
            Assert.AreEqual("C1", tableBuilder.CurrentState.Cursor.CellReference);
            Assert.AreEqual(0, tableBuilder.CurrentState.GlobalHeight);
            Assert.AreEqual(2, tableBuilder.CurrentState.GlobalWidth);
            Assert.AreEqual(0, tableBuilder.CurrentState.CurrentLayerHeight);
            Assert.AreEqual(1, tableBuilder.CurrentState.CurrentLayerStartRowIndex);
        }

        [Test]
        public void TableBuilderToNextLayerMovingTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var tableBuilder = new TableBuilder(table, new CellPosition("C4"));

            tableBuilder.RenderAtomicValue("1");
            tableBuilder.MoveToNextColumn();
            tableBuilder.RenderAtomicValue("2");
            tableBuilder.MoveToNextColumn();
            tableBuilder.MoveToNextLayer();
            Assert.AreEqual("C5", tableBuilder.CurrentState.Cursor.CellReference);
            Assert.AreEqual(5, tableBuilder.CurrentState.CurrentLayerStartRowIndex);
            Assert.AreEqual(1, tableBuilder.CurrentState.GlobalHeight);
            Assert.AreEqual(2, tableBuilder.CurrentState.GlobalWidth);
            Assert.AreEqual(0, tableBuilder.CurrentState.CurrentLayerHeight);
        }

        [Test]
        public void TableBuilderBigAwfulTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var tableBuilder = new TableBuilder(table, new CellPosition("A1"));

            tableBuilder.PushState() //depth = 1
                        .PushState() //depth = 2
                        .PushState() //depth = 3
                        .RenderAtomicValue("1")
                        .MoveToNextColumn()
                        .RenderAtomicValue("1")
                        .MoveToNextColumn()
                        .MoveToNextLayer()
                        .RenderAtomicValue("1")
                        .MoveToNextColumn()
                        .RenderAtomicValue("1")
                        .MoveToNextColumn()
                        .PopState() //depth = 2
                        .PushState() //depth = 3
                        .RenderAtomicValue("2")
                        .MoveToNextColumn()
                        .MoveToNextLayer()
                        .RenderAtomicValue("2")
                        .MoveToNextColumn()
                        .MoveToNextLayer()
                        .RenderAtomicValue("2")
                        .MoveToNextColumn()
                        .PopState() //depth = 2
                        .PopState() //depth = 1
                        .MoveToNextLayer()
                        .PushState() //depth = 2
                        .PushState() //depth = 3
                        .RenderAtomicValue("3")
                        .MoveToNextColumn()
                        .RenderAtomicValue("3")
                        .MoveToNextColumn()
                        .RenderAtomicValue("3")
                        .MoveToNextColumn()
                        .PopState() //depth = 2
                        .PushState() //depth = 3
                        .RenderAtomicValue("4")
                        .MoveToNextColumn()
                        .MoveToNextLayer()
                        .RenderAtomicValue("4")
                        .MoveToNextColumn()
                        .MoveToNextLayer()
                        .RenderAtomicValue("4")
                        .MoveToNextColumn()
                        .PopState() //depth = 2
                        .PopState() //depth = 1
                        .PushState() //depth = 2
                        .PushState() //depth = 3
                        .RenderAtomicValue("5")
                        .MoveToNextColumn()
                        .MoveToNextLayer()
                        .RenderAtomicValue("5")
                        .MoveToNextColumn()
                        .PopState() //depth = 2
                        .PushState() //depth = 3
                        .RenderAtomicValue("6")
                        .MoveToNextColumn()
                        .RenderAtomicValue("6")
                        .MoveToNextColumn()
                        .PopState() //depth = 2
                        .PopState() //depth = 1
                        .MoveToNextLayer()
                        .RenderAtomicValue("7")
                        .MoveToNextColumn()
                        .PushState(); //depth = 0

            Assert.AreEqual(table.GetCell(new CellPosition("A1")).StringValue, "1");
            Assert.AreEqual(table.GetCell(new CellPosition("B1")).StringValue, "1");
            Assert.AreEqual(table.GetCell(new CellPosition("A2")).StringValue, "1");
            Assert.AreEqual(table.GetCell(new CellPosition("B2")).StringValue, "1");
            Assert.AreEqual(table.GetCell(new CellPosition("C1")).StringValue, "2");
            Assert.AreEqual(table.GetCell(new CellPosition("C2")).StringValue, "2");
            Assert.AreEqual(table.GetCell(new CellPosition("C3")).StringValue, "2");
            Assert.AreEqual(table.GetCell(new CellPosition("A4")).StringValue, "3");
            Assert.AreEqual(table.GetCell(new CellPosition("B4")).StringValue, "3");
            Assert.AreEqual(table.GetCell(new CellPosition("C4")).StringValue, "3");
            Assert.AreEqual(table.GetCell(new CellPosition("D4")).StringValue, "4");
            Assert.AreEqual(table.GetCell(new CellPosition("D5")).StringValue, "4");
            Assert.AreEqual(table.GetCell(new CellPosition("D6")).StringValue, "4");
            Assert.AreEqual(table.GetCell(new CellPosition("E4")).StringValue, "5");
            Assert.AreEqual(table.GetCell(new CellPosition("E5")).StringValue, "5");
            Assert.AreEqual(table.GetCell(new CellPosition("F4")).StringValue, "6");
            Assert.AreEqual(table.GetCell(new CellPosition("G4")).StringValue, "6");
            Assert.AreEqual(table.GetCell(new CellPosition("A7")).StringValue, "7");
        }
    }
}