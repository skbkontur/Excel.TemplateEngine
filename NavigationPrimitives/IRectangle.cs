namespace SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives
{
    public interface IRectangle
    {
        ICellPosition UpperLeft { get; }
        ICellPosition LowerRight { get; }
        IObjectSize Size { get; }
        bool IsIntersects(IRectangle rect);
    }
}