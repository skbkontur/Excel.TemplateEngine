using System;
using System.Collections.Generic;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter
{
    public interface ITemplateEngine
    {
        void Render<TModel>([NotNull] ITableBuilder tableBuilder, [NotNull] TModel model);

        (TModel model, Dictionary<string, string> mappingForErrors) Parse<TModel>([NotNull] ITableParser tableParser)
            where TModel : new();
    }
}