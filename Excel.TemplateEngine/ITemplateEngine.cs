using System.Collections.Generic;

using Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using Excel.TemplateEngine.ObjectPrinting.TableParser;

using JetBrains.Annotations;

namespace Excel.TemplateEngine
{
    public interface ITemplateEngine
    {
        void Render<TModel>([NotNull] ITableBuilder tableBuilder, [NotNull] TModel model);

        (TModel model, Dictionary<string, string> mappingForErrors) Parse<TModel>([NotNull] ITableParser tableParser)
            where TModel : new();
    }
}