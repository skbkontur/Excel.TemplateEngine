using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public class IntRenderer : IRenderer
    {
        public void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template)
        {
            if (!(model is int))
                throw new ArgumentException("model is not int");

            var intToRender = (int)model;
            tableBuilder.RenderAtomicValue(intToRender);
            tableBuilder.SetCurrentStyle();
            tableBuilder.MoveToNextColumn();
        }
    }
}