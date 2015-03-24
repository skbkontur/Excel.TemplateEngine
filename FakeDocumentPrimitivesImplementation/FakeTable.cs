using System;
using System.Collections.Generic;
using System.Linq;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.FakeDocumentPrimitivesImplementation
{
    public class FakeTable : ITable
    {
        public FakeTable(int width, int height)
        {
            this.width = width;
            this.height = height;

            cells = JaggedArrayHelper.Instance.CreateJaggedArray<ICell[][]>(height, width);
            MergedCells = new List<Tuple<ICellPosition, ICellPosition>>();
        }

        public ICell GetCell(ICellPosition position)
        {
            return OutOfBounds(position) ? null : cells[position.RowIndex - 1][position.ColumnIndex - 1];
        }

        public ICell InsertCell(ICellPosition position)
        {
            if(OutOfBounds(position))
                return null;

            var newCell = new FakeCell(position)
                {
                    StyleId = position.CellReference
                };
            cells[position.RowIndex - 1][position.ColumnIndex - 1] = newCell;
            return newCell;
        }

        public IEnumerable<ICell> SearchCellByText(string text)
        {
            return from row in cells
                   from cell in row
                   where
                       cell != null && (cell.StringValue + "").Contains(text)
                   select cell;
        }

        public ITablePart GetTablePart(ICellPosition upperLeft, ICellPosition lowerRight)
        {
            if(OutOfBounds(upperLeft) || OutOfBounds(lowerRight))
                return null;

            var subTableSize = lowerRight.Subtract(upperLeft).Add(new ObjectSize(1, 1));
            var subTable = JaggedArrayHelper.Instance.CreateJaggedArray<ICell[][]>(subTableSize.Height, subTableSize.Width);
            for(var y = 0; y < subTableSize.Height; ++y)
            {
                for(var x = 0; x < subTableSize.Width; ++x)
                {
                    var sourceCell = cells[upperLeft.RowIndex + y - 1][upperLeft.ColumnIndex + x - 1];
                    subTable[y][x] = sourceCell ?? new FakeCell(new CellPosition(y + 1, x + 1));
                }
            }
            return new FakeTablePart(subTable);
        }

        public IEnumerable<IColumn> Columns { get { return Enumerable.Range(0, cells[0].Count()).Select(index => new FakeColumn {Index = index}); } }

        public void MergeCells(ICellPosition upperLeft, ICellPosition lowerRight)
        {
            MergedCells.Add(Tuple.Create(upperLeft, lowerRight));
        }

        public void ResizeColumn(int columnIndex, double columnWidth)
        {
        }

        public static FakeTable GenerateFromStringArray(string[][] template)
        {
            if(!CheckArrayDimentions(template))
                return null;
            var height = template.Count();
            var width = template[0].Count();

            var table = new FakeTable(width, height);
            for(var y = 0; y < height; ++y)
            {
                for(var x = 0; x < width; ++x)
                {
                    var cell = table.InsertCell(new CellPosition(y + 1, x + 1));
                    cell.StringValue = template[y][x];
                }
            }
            return table;
        }

        private static bool CheckArrayDimentions(string[][] array)
        {
            if(array == null || !array.Any() || !array[0].Any())
                return false;
            var width = array[0].Count();
            return array.All(row => row.Any() && row.Count() == width);
        }

        private bool OutOfBounds(ICellPosition position)
        {
            return position.RowIndex < 1 || position.RowIndex > height ||
                   position.ColumnIndex < 1 || position.ColumnIndex > width;
        }

        public List<Tuple<ICellPosition, ICellPosition>> MergedCells { get; private set; }

        private readonly int width;
        private readonly int height;
        private readonly ICell[][] cells;
    }
}