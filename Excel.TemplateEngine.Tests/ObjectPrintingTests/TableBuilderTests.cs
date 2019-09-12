using System.Linq;

using Excel.TemplateEngine.Tests.ObjectPrintingTests.FakeDocumentPrimitivesImplementation;

using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator;

using Vostok.Logging.Console;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests
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

            var tableBuilder = new TableBuilder(table, new TableNavigator(new CellPosition("A1"), logger));

            tableBuilder.RenderAtomicValue("Test");
            tableBuilder.MoveToNextColumn();
            table.GetCell(new CellPosition("A1")).StringValue.Should().Be("Test");
            tableBuilder.CurrentState.Cursor.CellReference.Should().Be("B1");
            tableBuilder.CurrentState.GlobalHeight.Should().Be(0);
            tableBuilder.CurrentState.GlobalWidth.Should().Be(1);
            tableBuilder.CurrentState.CurrentLayerHeight.Should().Be(0);
            tableBuilder.CurrentState.CurrentLayerStartRowIndex.Should().Be(1);

            tableBuilder.RenderAtomicValue("tseT");
            tableBuilder.MoveToNextColumn();
            table.GetCell(new CellPosition("B1")).StringValue.Should().Be("tseT");
            tableBuilder.CurrentState.Cursor.CellReference.Should().Be("C1");
            tableBuilder.CurrentState.GlobalHeight.Should().Be(0);
            tableBuilder.CurrentState.GlobalWidth.Should().Be(2);
            tableBuilder.CurrentState.CurrentLayerHeight.Should().Be(0);
            tableBuilder.CurrentState.CurrentLayerStartRowIndex.Should().Be(1);
        }

        [Test]
        public void TableBuilderToNextLayerMovingTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var tableBuilder = new TableBuilder(table, new TableNavigator(new CellPosition("C4"), logger));

            tableBuilder.RenderAtomicValue("1");
            tableBuilder.MoveToNextColumn();
            tableBuilder.RenderAtomicValue("2");
            tableBuilder.MoveToNextColumn();
            tableBuilder.MoveToNextLayer();
            tableBuilder.CurrentState.Cursor.CellReference.Should().Be("C5");
            tableBuilder.CurrentState.CurrentLayerStartRowIndex.Should().Be(5);
            tableBuilder.CurrentState.GlobalHeight.Should().Be(1);
            tableBuilder.CurrentState.GlobalWidth.Should().Be(2);
            tableBuilder.CurrentState.CurrentLayerHeight.Should().Be(0);
        }

        [Test]
        public void TableBuilderBigAwfulTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var tableBuilder = new TableBuilder(table, new TableNavigator(new CellPosition("A1"), logger));

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

            table.GetCell(new CellPosition("A1")).StringValue.Should().Be("1");
            table.GetCell(new CellPosition("B1")).StringValue.Should().Be("1");
            table.GetCell(new CellPosition("A2")).StringValue.Should().Be("1");
            table.GetCell(new CellPosition("B2")).StringValue.Should().Be("1");
            table.GetCell(new CellPosition("C1")).StringValue.Should().Be("2");
            table.GetCell(new CellPosition("C2")).StringValue.Should().Be("2");
            table.GetCell(new CellPosition("C3")).StringValue.Should().Be("2");
            table.GetCell(new CellPosition("A4")).StringValue.Should().Be("3");
            table.GetCell(new CellPosition("B4")).StringValue.Should().Be("3");
            table.GetCell(new CellPosition("C4")).StringValue.Should().Be("3");
            table.GetCell(new CellPosition("D4")).StringValue.Should().Be("4");
            table.GetCell(new CellPosition("D5")).StringValue.Should().Be("4");
            table.GetCell(new CellPosition("D6")).StringValue.Should().Be("4");
            table.GetCell(new CellPosition("E4")).StringValue.Should().Be("5");
            table.GetCell(new CellPosition("E5")).StringValue.Should().Be("5");
            table.GetCell(new CellPosition("F4")).StringValue.Should().Be("6");
            table.GetCell(new CellPosition("G4")).StringValue.Should().Be("6");
            table.GetCell(new CellPosition("A7")).StringValue.Should().Be("7");
        }

        [Test]
        public void TableBuilderBigAwfulTestWithMergedCells()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var tableBuilder = new TableBuilder(table, new TableNavigator(new CellPosition("A1"), logger));

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
                        .MergeCells(new Rectangle(new CellPosition("A1"), new CellPosition("B2")))
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
                        .MergeCells(new Rectangle(new CellPosition("A2"), new CellPosition("A3")))
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

            var mergedCells = table.MergedCells.ToArray();

            mergedCells.Length.Should().Be(2);
            mergedCells[0].UpperLeft.CellReference.Should().Be("A1");
            mergedCells[0].LowerRight.CellReference.Should().Be("B2");
            mergedCells[1].UpperLeft.CellReference.Should().Be("D5");
            mergedCells[1].LowerRight.CellReference.Should().Be("D6");
        }

        private readonly ConsoleLog logger = new ConsoleLog();
    }
}