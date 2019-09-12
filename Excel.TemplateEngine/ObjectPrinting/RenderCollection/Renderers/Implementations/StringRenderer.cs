using System;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers.Implementations
{
    internal class StringRenderer : IRenderer
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