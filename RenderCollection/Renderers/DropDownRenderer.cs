using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public class DropDownRenderer : IFormControlRenderer
    {
        public void Render(ITableBuilder tableBuilder, string name, object model)
        {
            if (!(model is string stringToRender))
                throw new ArgumentException("model is not string");

            tableBuilder.RenderDropDownValue(name, stringToRender);
        }
    }
}