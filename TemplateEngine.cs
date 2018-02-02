using System.Collections.Generic;

using log4net;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter
{
    public class TemplateEngine : ITemplateEngine
    {
        public TemplateEngine(ITable templateTable)
        {
            formControlsInfo = templateTable.GetFormControlsInfo();
            templateCollection = new TemplateCollection(templateTable);
            rendererCollection = new RendererCollection(templateCollection);
            parserCollection = new ParserCollection();
        }

        public void Render(ITableBuilder tableBuilder, object model)
        {
            var renderingTemplate = templateCollection.GetTemplate(rootTemplateName);
            if(renderingTemplate == null)
            {
                RenderError(tableBuilder);
                return;
            }
            tableBuilder.AddFormControlInfos(formControlsInfo);
            var render = rendererCollection.GetRenderer(model.GetType());
            render.Render(tableBuilder, model, renderingTemplate);
        }

        public (TModel model, Dictionary<string, string> mappingForErrors) Parse<TModel>(ITableParser tableParser)
            where TModel : new()
        {
            var renderingTemplate = templateCollection.GetTemplate(rootTemplateName);

            if(renderingTemplate == null)
                throw new InvalidProgramStateException($"Template with name {rootTemplateName} not found in xlsx");

            var parser = parserCollection.GetClassParser(typeof(TModel));
            var fieldsMappingForErrors = new Dictionary<string, string>();
            return (parser.Parse<TModel>(tableParser, renderingTemplate, (name, value) => fieldsMappingForErrors.Add(name, value)), fieldsMappingForErrors);
        }

        private void RenderError(ITableBuilder tableBuilder)
        {
            tableBuilder.RenderAtomicValue("Error: Root template description not found!");
            logger.Error("Excel document generation failed: root template description not found.");
        }

        private const string rootTemplateName = "RootTemplate";
        private readonly ITemplateCollection templateCollection;
        private readonly IRendererCollection rendererCollection;
        private readonly ILog logger = LogManager.GetLogger(typeof(TemplateEngine));
        private readonly IParserCollection parserCollection;
        private readonly IFormControls formControlsInfo;
    }
}