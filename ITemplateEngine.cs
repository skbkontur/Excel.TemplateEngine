using System.Collections.Generic;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter
{
    public interface ITemplateEngine
    {
        void Render(ITableBuilder tableBuilder, object model);
        (TModel model, Dictionary<string, string> mappingForErrors) Parse<TModel>(ITableParser tableParser);
    }
}