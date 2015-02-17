using System.Collections.Generic;
using System.Linq;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ExcelDocumentPrimitivesImplementation
{
    public class ExcelTable : ITable
    {
        public ExcelTable(IExcelWorksheet excelWotksheet)
        {
            internalTable = excelWotksheet;
        }

        public ICell GetCell(ICellPosition position)
        {
            var internalCell = internalTable.GetCell(new ExcelCellIndex(position.CellReference));
            return internalCell == null ? null : new ExcelCell(internalCell);
        }

        public ICell InsertCell(ICellPosition position)
        {
            return new ExcelCell(internalTable.InsertCell(new ExcelCellIndex(position.CellReference)));
        }

        public IEnumerable<ICell> SearchCellByText(string text)
        {
            return internalTable.SearchCellsByText(text).Select(cell => new ExcelCell(cell));
        }

        public ITablePart GetTablePart(ICellPosition upperLeft, ICellPosition lowerRight)
        {
            var excelReferenceToCell = internalTable
                .GetSortedCellsInRange(new ExcelCellIndex(upperLeft.CellReference), new ExcelCellIndex(lowerRight.CellReference))
                .Select(cell => new ExcelCell(cell))
                .ToDictionary(cell => cell.CellPosition.CellReference);

            var subTableSize = lowerRight.Subtract(upperLeft).Add(new ObjectSize(1, 1));
            var subTable = JaggedArrayHelper.Instance.CreateJaggedArray<ICell[][]>(subTableSize.Height, subTableSize.Width);
            for(var x = upperLeft.ColumnIndex; x <= lowerRight.ColumnIndex; ++x)
            {
                for(var y = upperLeft.RowIndex; y <= lowerRight.RowIndex; ++y)
                {
                    ExcelCell cell;
                    if(!excelReferenceToCell.TryGetValue(new CellPosition(y, x).CellReference, out cell))
                        cell = new ExcelCell(internalTable.InsertCell(new ExcelCellIndex(y, x)));

                    subTable[y - upperLeft.RowIndex][x - upperLeft.ColumnIndex] = cell;
                }
            }

            return new ExcelTablePart(subTable);
        }

        private readonly IExcelWorksheet internalTable;
    }
}