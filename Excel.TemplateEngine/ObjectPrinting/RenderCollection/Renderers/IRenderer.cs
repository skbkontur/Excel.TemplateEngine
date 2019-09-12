using Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
{
    internal interface IRenderer
    {
        void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template);
    }
}