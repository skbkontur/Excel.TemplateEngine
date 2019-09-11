using Excel.TemplateEngine.Exceptions;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
{
    public class DropDownRenderer : IFormControlRenderer
    {
        public void Render([NotNull] ITableBuilder tableBuilder, [NotNull] string name, [NotNull] object model)
        {
            if (!(model is string stringToRender))
                throw new ExcelTemplateEngineException("model is not string");

            tableBuilder.RenderDropDownValue(name, stringToRender);
        }
    }
}