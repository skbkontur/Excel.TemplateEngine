using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter
{
    public interface ITemplateEngine
    {
        void Render(ITableBuilder tableBuilder, object model);
    }
}