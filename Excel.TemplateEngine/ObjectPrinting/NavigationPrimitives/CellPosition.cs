namespace SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives
{
    public class CellPosition : ICellPosition
    {
        public CellPosition(int rowIndex, int columnIndex)
        {
            internalCellIndex = new ExcelCellIndex(rowIndex, columnIndex);
        }

        public CellPosition(string cellReference)
        {
            internalCellIndex = new ExcelCellIndex(cellReference);
        }

        public CellPosition(ExcelCellIndex excelCellIndex)
        {
            internalCellIndex = excelCellIndex;
        }

        public ICellPosition Add(IObjectSize other)
        {
            return new CellPosition(RowIndex + other.Height, ColumnIndex + other.Width);
        }

        public IObjectSize Subtract(ICellPosition other)
        {
            return new ObjectSize(ColumnIndex - other.ColumnIndex, RowIndex - other.RowIndex);
        }

        public int RowIndex { get { return internalCellIndex.RowIndex; } }
        public int ColumnIndex { get { return internalCellIndex.ColumnIndex; } }
        public string CellReference { get { return internalCellIndex.CellReference; } }

        private readonly ExcelCellIndex internalCellIndex;
    }
}