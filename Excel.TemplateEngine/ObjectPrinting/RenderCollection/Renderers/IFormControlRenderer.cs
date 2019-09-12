using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
{
    internal interface IFormControlRenderer
    {
        void Render([NotNull] ITableBuilder tableBuilder, [NotNull] string name, [NotNull] object model);
    }
}