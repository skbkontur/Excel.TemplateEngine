using System;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

#nullable enable

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers;

internal interface IClassParser
{
    TModel Parse<TModel>(ITableParser tableParser, RenderingTemplate template, Action<string, string> addFieldMapping)
        where TModel : new();

    void Parse<TModel>(ITableParser tableParser, RenderingTemplate template, Action<string, string> addFieldMapping, ref TModel model)
        where TModel : new();
}