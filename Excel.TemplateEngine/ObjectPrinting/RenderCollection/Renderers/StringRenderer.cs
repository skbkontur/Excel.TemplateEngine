using System;

using Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
{
    public class StringRenderer : IRenderer
    {
        public void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template)
        {
            if (!(model is string stringToRender))
                throw new ArgumentException("model is not string");

            tableBuilder.RenderAtomicValue(stringToRender);
            tableBuilder.SetCurrentStyle();
            tableBuilder.MoveToNextColumn();
        }
    }
}