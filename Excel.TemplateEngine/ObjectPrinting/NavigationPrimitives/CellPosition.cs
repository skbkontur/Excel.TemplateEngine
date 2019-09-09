using Excel.TemplateEngine.FileGenerating.Implementation;

namespace Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives
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

        public int RowIndex => internalCellIndex.RowIndex;
        public int ColumnIndex => internalCellIndex.ColumnIndex;
        public string CellReference => internalCellIndex.CellReference;

        private readonly ExcelCellIndex internalCellIndex;
    }
}