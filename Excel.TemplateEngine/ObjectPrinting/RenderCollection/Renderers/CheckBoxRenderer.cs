using Excel.TemplateEngine.ObjectPrinting.TableBuilder;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
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