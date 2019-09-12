using Excel.TemplateEngine.ObjectPrinting.TableBuilder;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
{
    internal interface IFormControlRenderer
    {
        void Render([NotNull] ITableBuilder tableBuilder, [NotNull] string name, [NotNull] object model);
    }
}