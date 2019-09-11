using System.Linq;

using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.FakeDocumentPrimitivesImplementation;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

using FluentAssertions;

using NUnit.Framework;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests
{
    [TestFixture]
    public class FakeDocumentPrimitivesTests
    {
        [Test]
        public void FakeTableCellManipulationTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var position = new CellPosition(10, 10);
            var cell = table.InsertCell(position);
            cell.Should().NotBeNull();

            const string testValue = "Test Value";
            cell.StringValue = testValue;

            cell = table.GetCell(position);

            cell.Should().NotBeNull();
            cell.StringValue.Should().Be(testValue);
        }

        [Test]
        public void FakeTableCellInsertionTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var position = new CellPosition(10, 10);
            var cell = table.InsertCell(position);
            cell.Should().NotBeNull();
        }

        [Test]
        [TestCase(-1, 10)]
        [TestCase(10, -1)]
        [TestCase(100, 10)]
        [TestCase(10, 100)]
        [TestCase(-1, -1)]
        [TestCase(100, 100)]
        public void FakeTableWrongPositionCellInsertionTest(int wrongRowIndex, int wrongColumnIndex)
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var cell = table.InsertCell(new CellPosition(wrongRowIndex, wrongColumnIndex));
            cell.Should().BeNull();
        }

        [Test]
        public void FakeTableEmptyCellExtractionTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var cell = table.GetCell(new CellPosition(-1, -1));
            cell.Should().BeNull();

            cell = table.GetCell(new CellPosition(3, 3));
            cell.Should().BeNull();
        }

        [Test]
        public void FakeTableCellSearchTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var positions = new[] {new CellPosition(1, 1), new CellPosition(40, 20), new CellPosition(10, 10)};
            var stringValues = new[] {"Test Value", "Test Test", "Another text"};

            for (var i = 0; i < positions.Count(); ++i)
            {
                var cell = table.InsertCell(positions[i]);
                cell.StringValue = stringValues[i];
            }

            var cells = table.SearchCellByText("Test");
            var cellsArray = cells as ICell[] ?? cells.ToArray();

            cellsArray.Length.Should().Be(2);
            cellsArray[0].Should().NotBeNull();
            cellsArray[1].Should().NotBeNull();

            cellsArray[0].CellPosition.RowIndex.Should().Be(1);
            cellsArray[0].CellPosition.ColumnIndex.Should().Be(1);
            cellsArray[1].CellPosition.RowIndex.Should().Be(40);
            cellsArray[1].CellPosition.ColumnIndex.Should().Be(20);
        }

        [Test]
        public void FakeTablePartExtractionTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            for (var x = 0; x < width; ++x)
            {
                for (var y = 0; y < height; ++y)
                    table.InsertCell(new CellPosition(y + 1, x + 1));
            }

            var positions = new[] {new CellPosition(1, 1), new CellPosition(40, 20), new CellPosition(10, 10)};
            var stringValues = new[] {"Test Value", "Test Test", "Another text"};

            for (var i = 0; i < positions.Count(); ++i)
            {
                var cell = table.InsertCell(positions[i]);
                cell.StringValue = stringValues[i];
            }

            var tablePart = table.GetTablePart(new Rectangle(new CellPosition(9, 9), new CellPosition(40, 20)));

            tablePart.Cells.Count().Should().Be(32);

            foreach (var row in tablePart.Cells)
                row.Count().Should().Be(12);

            var targetRow = tablePart.Cells.FirstOrDefault(row => row.FirstOrDefault(cell => cell.StringValue == "Another text") != null);
            targetRow.Should().NotBeNull();
// ReSharper disable AssignNullToNotNullAttribute
            var targetCell = targetRow.FirstOrDefault(cell => cell.StringValue == "Another text");
// ReSharper restore AssignNullToNotNullAttribute

            targetCell.Should().NotBeNull();
// ReSharper disable PossibleNullReferenceException
            targetCell.CellPosition.RowIndex.Should().Be(10);
// ReSharper restore PossibleNullReferenceException
            targetCell.CellPosition.ColumnIndex.Should().Be(10);

            targetRow = tablePart.Cells.LastOrDefault();
            targetRow.Should().NotBeNull();
// ReSharper disable AssignNullToNotNullAttribute
            targetCell = tablePart.Cells.LastOrDefault().LastOrDefault();
            // ReSharper restore AssignNullToNotNullAttribute
            targetCell.Should().NotBeNull();

// ReSharper disable PossibleNullReferenceException
            targetCell.StringValue.Should().Be("Test Test");
// ReSharper restore PossibleNullReferenceException
        }

        [Test]
        public void FakeTableWithWrongPositionPartExtractionTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var tablePart = table.GetTablePart(new Rectangle(new CellPosition(-1, 0), new CellPosition(1, 1)));

            tablePart.Should().BeNull();
        }

        [Test]
        public void FromStringArrayFakeTableInitializationTest()
        {
            var template = new[]
                {
                    new[]
                        {
                            "Text", null, ""
                        },
                    new[]
                        {
                            "Value:RootModel:Root", "qwe", "Model:ABC:A1:SD123"
                        }
                };
            var table = FakeTable.GenerateFromStringArray(template);

            table.GetCell(new CellPosition("A1")).StringValue.Should().Be("Text");
            table.GetCell(new CellPosition("B1")).StringValue.Should().BeNull();
            table.GetCell(new CellPosition(2, 3)).StringValue.Should().Be("Model:ABC:A1:SD123");
        }
    }
}