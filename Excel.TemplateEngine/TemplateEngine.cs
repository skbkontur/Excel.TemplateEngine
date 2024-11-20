using System;
using System.Collections.Generic;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

using Vostok.Logging.Abstractions;

#nullable enable
namespace SkbKontur.Excel.TemplateEngine;

public class TemplateEngine : ITemplateEngine
{
    public TemplateEngine(ITable templateTable, ILog logger)
    {
        this.templateTable = templateTable;
        templateCollection = new TemplateCollection(templateTable);
        rendererCollection = new RendererCollection(templateCollection);
        parserCollection = new ParserCollection(logger.ForContext("ExcelObjectPrinter"));
    }

    public void Render<TModel>(ITableBuilder tableBuilder, TModel model) where TModel : notnull
    {
        var renderingTemplate = templateCollection.GetTemplate(rootTemplateName)
                                ?? throw new InvalidOperationException($"Template with name {rootTemplateName} not found in xlsx");
        tableBuilder.CopyFormControlsFrom(templateTable);
        tableBuilder.CopyDataValidationsFrom(templateTable);
        tableBuilder.CopyWorksheetExtensionListFrom(templateTable); // WorksheetExtensionList contains info about data validations with ranges from other sheets, so copying it to support them.
        tableBuilder.CopyCommentsFrom(templateTable);
        var render = rendererCollection.GetRenderer(model.GetType());
        render.Render(tableBuilder, model, renderingTemplate);
    }

    public (TModel model, Dictionary<string, string> mappingForErrors) Parse<TModel>(ITableParser tableParser)
        where TModel : new()
    {
        var renderingTemplate = templateCollection.GetTemplate(rootTemplateName)
                                ?? throw new InvalidOperationException($"Template with name {rootTemplateName} not found in xlsx");
        var parser = parserCollection.GetClassParser();
        var fieldsMappingForErrors = new Dictionary<string, string>();
        return (model : parser.Parse<TModel>(tableParser, renderingTemplate, (name, value) => fieldsMappingForErrors.Add(name, value)), mappingForErrors : fieldsMappingForErrors);
    }

    public void Parse<TModel>(ITableParser tableParser, Action<string, string> mappingForErrors, ref TModel model)
        where TModel : new()
    {
        var renderingTemplate = templateCollection.GetTemplate(rootTemplateName)
                                ?? throw new InvalidOperationException($"Template with name {rootTemplateName} not found in xlsx");
        var parser = parserCollection.GetClassParser();
        parser.Parse(tableParser, renderingTemplate, mappingForErrors, ref model);
    }

    /// <summary>
    ///     Parse only separate cell values and List<> enumerations without and size limitations.
    /// </summary>
    /// <typeparam name="TModel">Class to parse.</typeparam>
    /// <param name="lazyTableReader">LazyTableReader of target xlsx file.</param>
    /// <param name="readerOffset">Target file offset relative to a template.</param>
    /// <param name="formulaEvaluator">Target file formula evaluator.</param>
    public TModel LazyParse<TModel>(LazyTableReader lazyTableReader, ObjectSize? readerOffset = null, IFormulaEvaluator? formulaEvaluator = null)
        where TModel : new()
    {
        readerOffset ??= new ObjectSize(0, 0);

        var renderingTemplate = templateCollection.GetTemplate(rootTemplateName) ??
                                throw new InvalidOperationException($"Template with name {rootTemplateName} not found in xlsx");
        var parser = parserCollection.GetLazyClassParser();
        return parser.Parse<TModel>(lazyTableReader, renderingTemplate, readerOffset, formulaEvaluator);
    }

    private const string rootTemplateName = "RootTemplate";

    private readonly ITable templateTable;

    private readonly ITemplateCollection templateCollection;

    private readonly IRendererCollection rendererCollection;

    private readonly IParserCollection parserCollection;
}