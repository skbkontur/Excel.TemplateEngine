using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers
{
    internal interface IRenderer
    {
        void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template);
    }
}