using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.Exceptions;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers.Implementations
{
    internal class DropDownRenderer : IFormControlRenderer
    {
        public void Render([NotNull] ITableBuilder tableBuilder, [NotNull] string name, [NotNull] object model)
        {
            if (!(model is string stringToRender))
                throw new ExcelTemplateEngineException("model is not string");

            tableBuilder.RenderDropDownValue(name, stringToRender);
        }
    }
}