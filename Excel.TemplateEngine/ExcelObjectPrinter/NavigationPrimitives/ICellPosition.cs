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

    public static class CellPositionExtentions
    {
        public static ICellPosition ToRelativeCoordinates(this ICellPosition globalCoordinates, ICellPosition newOrigin)
        {
            return new CellPosition(1, 1).Add(globalCoordinates.Subtract(newOrigin));
        }

        public static ICellPosition ToGlobalCoordinates(this ICellPosition relativeCoordinates, ICellPosition originCoordinates)
        {
            return relativeCoordinates.Add(originCoordinates.Subtract(new CellPosition(1, 1)));
        }
    }
}