using System;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers.Implementations
{
    internal class DecimalRenderer : IRenderer
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