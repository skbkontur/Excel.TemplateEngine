using System.Collections.Generic;

using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.ParseCollection;
using Excel.TemplateEngine.ObjectPrinting.RenderCollection;
using Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using Excel.TemplateEngine.ObjectPrinting.TableParser;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting
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
                                    ?? throw new ExcelEngineException($"Template with name {rootTemplateName} not found in xlsx");
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
                                    ?? throw new ExcelEngineException($"Template with name {rootTemplateName} not found in xlsx");
            var parser = parserCollection.GetClassParser();
            var fieldsMappingForErrors = new Dictionary<string, string>();
            return (model : parser.Parse<TModel>(tableParser, renderingTemplate, (name, value) => fieldsMappingForErrors.Add(name, value)), mappingForErrors : fieldsMappingForErrors);
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