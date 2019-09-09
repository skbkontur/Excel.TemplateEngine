using System;

using Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
{
    public class DecimalRenderer : IRenderer
    {
        public void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template)
        {
            if (!(model is decimal))
                throw new ArgumentException("model is not decimal");

            var decimalToRender = (decimal)model;
            tableBuilder.RenderAtomicValue(decimalToRender);
            tableBuilder.SetCurrentStyle();
            tableBuilder.MoveToNextColumn();
        }
    }
}