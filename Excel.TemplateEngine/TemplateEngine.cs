using System;
using System.Collections.Generic;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine
{
    public class TemplateEngine : ITemplateEngine
    {
        public TemplateEngine([NotNull] ITable templateTable, [NotNull] ILog logger)
        {
            this.templateTable = templateTable;
            templateCollection = new TemplateCollection(templateTable);
            rendererCollection = new RendererCollection(templateCollection);
            parserCollection = new ParserCollection(logger.ForContext("ExcelObjectPrinter"));
        }

        public void Render<TModel>([NotNull] ITableBuilder tableBuilder, [NotNull] TModel model)
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

        public (TModel model, Dictionary<string, string> mappingForErrors) Parse<TModel>([NotNull] ITableParser tableParser)
            where TModel : new()
        {
            var renderingTemplate = templateCollection.GetTemplate(rootTemplateName)
                                    ?? throw new InvalidOperationException($"Template with name {rootTemplateName} not found in xlsx");
            var parser = parserCollection.GetClassParser();
            var fieldsMappingForErrors = new Dictionary<string, string>();
            return (model : parser.Parse<TModel>(tableParser, renderingTemplate, (name, value) => fieldsMappingForErrors.Add(name, value)), mappingForErrors : fieldsMappingForErrors);
        }

        /// <summary>
        /// Parse only separate cell values and List<> enumerations without and size limitations.
        /// </summary>
        /// <typeparam name="TModel">Class to parse.</typeparam>
        /// <param name="lazyTableReader">LazyTableReader of target xlsx file.</param>
        public TModel LazyParse<TModel>([NotNull] LazyTableReader lazyTableReader)
            where TModel : new()
        {
            var renderingTemplate = templateCollection.GetTemplate(rootTemplateName) ??
                                    throw new InvalidOperationException($"Template with name {rootTemplateName} not found in xlsx");
            var parser = parserCollection.GetLazyClassParser();
            return parser.Parse<TModel>(lazyTableReader, renderingTemplate);
        }

        private const string rootTemplateName = "RootTemplate";

        [NotNull]
        private readonly ITable templateTable;

        [NotNull]
        private readonly ITemplateCollection templateCollection;

        [NotNull]
        private readonly IRendererCollection rendererCollection;

        [NotNull]
        private readonly IParserCollection parserCollection;
    }
}