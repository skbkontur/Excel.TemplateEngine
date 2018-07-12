using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public interface IFormControlRenderer
    {
        void Render([NotNull] ITableBuilder tableBuilder, [NotNull] string name, [NotNull] object model);
    }
}