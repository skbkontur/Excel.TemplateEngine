using System.Collections;
using System.Linq;

using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public class EnumerableRenderer : IRenderer
    {
        public EnumerableRenderer(IRendererCollection rendererCollection)
        {
            this.rendererCollection = rendererCollection;
        }

        public void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template)
        {
            var enumerableToRender = ((IEnumerable)model).Cast<object>().ToArray();

            for(var i = 0; i < enumerableToRender.Length; ++i)
            {
                var element = enumerableToRender[i];

                var normalizedElement = NormalizeElement(element);

                tableBuilder.PushState();

                var renderer = rendererCollection.GetRenderer(normalizedElement.GetType());
                renderer.Render(tableBuilder, normalizedElement, template);

                tableBuilder.PopState();

                if(i != enumerableToRender.Length - 1)
                    tableBuilder.MoveToNextLayer();
            }
        }

        private static object NormalizeElement(object element)
        {
            return element ?? "";
        }

        private readonly IRendererCollection rendererCollection;
    }
}