using System.Collections.Generic;

using Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace Excel.TemplateEngine.ObjectPrinting
{
    public interface ITemplateEngine
    {
        void Render<TModel>([NotNull] ITableBuilder tableBuilder, [NotNull] TModel model);

        (TModel model, Dictionary<string, string> mappingForErrors) Parse<TModel>([NotNull] ITableParser tableParser)
            where TModel : new();
    }
}