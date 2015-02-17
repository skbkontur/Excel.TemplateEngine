namespace SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives
{
    public interface ICellPosition
    {
        ICellPosition Add(IObjectSize other);
        IObjectSize Subtract(ICellPosition other);
        int RowIndex { get; }
        int ColumnIndex { get; }
        string CellReference { get; }
    }
}