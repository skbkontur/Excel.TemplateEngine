using System;
using System.Collections.Generic;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

#nullable enable
namespace SkbKontur.Excel.TemplateEngine;

public interface ITemplateEngine
{
    void Render<TModel>(ITableBuilder tableBuilder, TModel model) where TModel : notnull;

    (TModel model, Dictionary<string, string> mappingForErrors) Parse<TModel>(ITableParser tableParser)
        where TModel : new();

    void Parse<TModel>(ITableParser tableParser, Action<string, string> mappingForErrors, ref TModel model)
        where TModel : new();

    public TModel LazyParse<TModel>(LazyTableReader lazyTableReader, ObjectSize? readerOffset = null, IFormulaEvaluator? formulaEvaluator = null)
        where TModel : new();
}