using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder
{
    public interface IStyler
    {
        void ApplyStyle(ICell cell);
    }
}