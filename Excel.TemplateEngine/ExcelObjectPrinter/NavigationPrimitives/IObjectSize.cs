namespace SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives
{
    public interface IObjectSize
    {
        IObjectSize Add(IObjectSize other);
        IObjectSize Subtract(IObjectSize other);
        int Width { get; }
        int Height { get; }
    }
}