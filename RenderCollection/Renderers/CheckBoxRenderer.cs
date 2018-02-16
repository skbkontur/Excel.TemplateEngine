using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public class CheckBoxRenderer : IFormControlRenderer
    {
        public void Render(ITableBuilder tableBuilder, string name, object model)
        {
            if(!(model is bool boolToRender))
                throw new ArgumentException("model is not bool");

            tableBuilder.RenderCheckBoxValue(name, boolToRender);
        }
    }
}