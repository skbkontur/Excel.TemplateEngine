using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public class DropDownRenderer : IFormControlRenderer
    {
        public void Render([NotNull] ITableBuilder tableBuilder, [NotNull] string name, [NotNull] object model)
        {
            if (!(model is string stringToRender))
                throw new InvalidProgramStateException("model is not string");

            tableBuilder.RenderDropDownValue(name, stringToRender);
        }
    }
}