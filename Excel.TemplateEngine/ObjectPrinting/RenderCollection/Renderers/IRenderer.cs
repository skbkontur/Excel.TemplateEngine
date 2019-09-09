using Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
{
    public interface IRenderer
    {
        void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template);
    }
}