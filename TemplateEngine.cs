using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.PostBuildActions;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

using log4net;

namespace SKBKontur.Catalogue.ExcelObjectPrinter
{
    public class TemplateEngine : ITemplateEngine
    {
        public TemplateEngine(ITable templateTable)
        {
            templateCollection = new TemplateCollection(templateTable);
            rendererCollection = new RendererCollection(templateCollection);
            columnResizer = new ColumnResizer(templateTable);
        }

        public void Render(ITableBuilder tableBuilder, object model)
        {
            var renderingTemplate = templateCollection.GetTemplate(rootTemplateName);
            if(renderingTemplate == null)
            {
                RenderError(tableBuilder);
                return;
            }

            var render = rendererCollection.GetRenderer(model.GetType());
            render.Render(tableBuilder, model, renderingTemplate);

            columnResizer.ResizeColumns(tableBuilder);
        }

        private void RenderError(ITableBuilder tableBuilder)
        {
            tableBuilder.RenderAtomicValue("Error: Root template description not found!");
            logger.Error("Excel document generation failed: root template description not found.");
        }

        private readonly ITemplateCollection templateCollection;
        private readonly IRendererCollection rendererCollection;
        private readonly IColumnResizer columnResizer;
        private readonly ILog logger = LogManager.GetLogger(typeof(TemplateEngine));

        private const string rootTemplateName = "RootTemplate";
    }
}