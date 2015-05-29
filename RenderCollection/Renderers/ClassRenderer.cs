using System.Linq;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public class ClassRenderer : IRenderer
    {
        public ClassRenderer(ITemplateCollection templateCollection, IRendererCollection rendererCollection)
        {
            this.templateCollection = templateCollection;
            this.rendererCollection = rendererCollection;
        }

        public void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template)
        {
            var lastRow = template.Content.Cells.LastOrDefault();
            foreach(var row in template.Content.Cells)
            {
                foreach(var cell in row)
                {
                    tableBuilder.PushState(new Styler(cell));

                    var childModel = ExtractChildModel(model, cell);
                    var childTemplateName = ExtractTemplateName(cell);
                    var renderer = rendererCollection.GetRenderer(childModel.GetType());
                    renderer.Render(tableBuilder, childModel, templateCollection.GetTemplate(childTemplateName));

                    tableBuilder.PopState();
                }
                if(!row.Equals(lastRow))
                    tableBuilder.MoveToNextLayer();
            }
            MergeCells(tableBuilder, template);
        }

        private static void MergeCells(ITableBuilder tableBuilder, RenderingTemplate template)
        {
            foreach(var mergedCells in template.MergedCells)
            {
                tableBuilder.MergeCells(mergedCells);
            }
        }

        private static string ExtractTemplateName(ICell cell)
        {
            return TemplateDescriptionHelper.Instance.ExtractTemplateNameFromValueDescription(cell.StringValue);
        }

        private static object ExtractChildModel(object model, ICell cell)
        {
            var expression = cell.StringValue;
            if(!TemplateDescriptionHelper.Instance.IsCorrectValueDescription(expression))
                return expression ?? "";
            var result = ObjectPropertiesExtractor.Instance.ExtractChildObject(model, expression);
            return result ?? "";
        }

        private readonly ITemplateCollection templateCollection;
        private readonly IRendererCollection rendererCollection;
    }
}