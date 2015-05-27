using System.Collections.Generic;

using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces
{
    public interface ITable
    {
        ICell GetCell(ICellPosition position);
        ICell InsertCell(ICellPosition position);
        IEnumerable<ICell> SearchCellByText(string text);
        ITablePart GetTablePart(IRectangle rectangle);
        void ResizeColumn(int columnIndex, double width);
        IEnumerable<IRectangle> MergedCells { get; }
        IEnumerable<IColumn> Columns { get; }
        void MergeCells(IRectangle rectangle);
    }
}