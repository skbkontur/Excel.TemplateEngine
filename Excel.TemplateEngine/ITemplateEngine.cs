using System.Collections.Generic;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace SkbKontur.Excel.TemplateEngine
{
    public interface ITemplateEngine
    {
        void Render<TModel>([NotNull] ITableBuilder tableBuilder, [NotNull] TModel model);

        (TModel model, Dictionary<string, string> mappingForErrors) Parse<TModel>([NotNull] ITableParser tableParser)
            where TModel : new();

        public TModel LazyParse<TModel>([NotNull] LazyTableReader lazyTableReader)
            where TModel : new();
    }
}