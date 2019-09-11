using System;

using Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using Excel.TemplateEngine.ObjectPrinting.TableParser;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    public interface IClassParser
    {
        [NotNull]
        TModel Parse<TModel>([NotNull] ITableParser tableParser, [NotNull] RenderingTemplate template, Action<string, string> addFieldMapping)
            where TModel : new();
    }
}