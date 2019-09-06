using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public class CheckBoxRenderer : IFormControlRenderer
    {
        public void Render([NotNull] ITableBuilder tableBuilder, [NotNull] string name, [NotNull] object model)
        {
            if (!(model is bool boolToRender))
                throw new InvalidProgramStateException("model is not bool");

            tableBuilder.RenderCheckBoxValue(name, boolToRender);
        }
    }
}