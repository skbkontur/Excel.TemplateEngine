using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator
{
    public interface IStyle
    {
        void ApplyTo(ICell cell);
    }
}