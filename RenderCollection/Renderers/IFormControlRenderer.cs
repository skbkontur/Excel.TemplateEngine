using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public interface IFormControlRenderer
    {
        void Render(ITableBuilder tableBuilder, string name, object model);
    }
}