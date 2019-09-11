using System.Linq;

using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.FakeDocumentPrimitivesImplementation;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

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
            Assert.AreNotEqual(null, cell);

            const string testValue = "Test Value";
            cell.StringValue = testValue;

            cell = table.GetCell(position);

            Assert.AreNotEqual(null, cell);
            Assert.AreEqual(testValue, cell.StringValue);
        }

        [Test]
        public void FakeTableCellInsertionTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var position = new CellPosition(10, 10);
            var cell = table.InsertCell(position);
            Assert.AreNotEqual(null, cell);
        }

        [Test]
        public void FakeTableWrongPositionCellInsertionTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var cell = table.InsertCell(new CellPosition(-1, 10));
            Assert.AreEqual(null, cell);

            cell = table.InsertCell(new CellPosition(10, -1));
            Assert.AreEqual(null, cell);

            cell = table.InsertCell(new CellPosition(100, 10));
            Assert.AreEqual(null, cell);

            cell = table.InsertCell(new CellPosition(10, 100));
            Assert.AreEqual(null, cell);

            cell = table.InsertCell(new CellPosition(-1, -1));
            Assert.AreEqual(null, cell);

            cell = table.InsertCell(new CellPosition(100, 100));
            Assert.AreEqual(null, cell);
        }

        [Test]
        public void FakeTableEmptyCellExtractionTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var cell = table.GetCell(new CellPosition(-1, -1));
            Assert.AreEqual(null, cell);

            cell = table.GetCell(new CellPosition(3, 3));
            Assert.AreEqual(null, cell);
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

            Assert.AreEqual(2, cellsArray.Count());
            Assert.AreNotEqual(null, cellsArray[0]);
            Assert.AreNotEqual(null, cellsArray[1]);

            Assert.AreEqual(1, cellsArray[0].CellPosition.RowIndex);
            Assert.AreEqual(1, cellsArray[0].CellPosition.ColumnIndex);
            Assert.AreEqual(40, cellsArray[1].CellPosition.RowIndex);
            Assert.AreEqual(20, cellsArray[1].CellPosition.ColumnIndex);
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

            Assert.AreEqual(32, tablePart.Cells.Count());

            foreach (var row in tablePart.Cells)
                Assert.AreEqual(12, row.Count());

            var targetRow = tablePart.Cells.FirstOrDefault(row => row.FirstOrDefault(cell => cell.StringValue == "Another text") != null);
            Assert.AreNotEqual(null, targetRow);
// ReSharper disable AssignNullToNotNullAttribute
            var targetCell = targetRow.FirstOrDefault(cell => cell.StringValue == "Another text");
// ReSharper restore AssignNullToNotNullAttribute

            Assert.AreNotEqual(null, targetCell);
// ReSharper disable PossibleNullReferenceException
            Assert.AreEqual(10, targetCell.CellPosition.RowIndex);
// ReSharper restore PossibleNullReferenceException
            Assert.AreEqual(10, targetCell.CellPosition.ColumnIndex);

            targetRow = tablePart.Cells.LastOrDefault();
            Assert.AreNotEqual(null, targetRow);

// ReSharper disable AssignNullToNotNullAttribute
            targetCell = tablePart.Cells.LastOrDefault().LastOrDefault();
// ReSharper restore AssignNullToNotNullAttribute
            Assert.AreNotEqual(null, targetCell);

// ReSharper disable PossibleNullReferenceException
            Assert.AreEqual(targetCell.StringValue, "Test Test");
// ReSharper restore PossibleNullReferenceException
        }

        [Test]
        public void FakeTableWithWrongPositionPartExtractionTest()
        {
            const int width = 20;
            const int height = 40;
            var table = new FakeTable(width, height);

            var tablePart = table.GetTablePart(new Rectangle(new CellPosition(-1, 0), new CellPosition(1, 1)));

            Assert.AreEqual(null, tablePart);
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

            Assert.AreEqual("Text", table.GetCell(new CellPosition("A1")).StringValue);
            Assert.AreEqual(null, table.GetCell(new CellPosition("B1")).StringValue);
            Assert.AreEqual("Model:ABC:A1:SD123", table.GetCell(new CellPosition(2, 3)).StringValue);
        }
    }
}