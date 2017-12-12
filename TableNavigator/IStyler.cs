using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator
{
    public interface IStyler
    {
        void ApplyStyle(ICell cell);
    }
}