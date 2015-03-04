using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public class DoubleRenderer : IRenderer
    {
        public void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template)
        {
            var doubleToRender = (double)model;
            tableBuilder.RenderAtomicValue(doubleToRender);
            tableBuilder.SetCurrentStyle();
            tableBuilder.MoveToNextColumn();
        }
    }
}