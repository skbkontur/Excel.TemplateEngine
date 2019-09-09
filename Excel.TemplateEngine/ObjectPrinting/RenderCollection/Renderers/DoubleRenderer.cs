using System;

using Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
{
    public class DoubleRenderer : IRenderer
    {
        public void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template)
        {
            if (!(model is double))
                throw new ArgumentException("model is not double");

            var doubleToRender = (double)model;
            tableBuilder.RenderAtomicValue(doubleToRender);
            tableBuilder.SetCurrentStyle();
            tableBuilder.MoveToNextColumn();
        }
    }
}