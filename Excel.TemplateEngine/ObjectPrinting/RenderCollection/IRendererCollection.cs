using System;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection
{
    internal interface IRendererCollection
    {
        IRenderer GetRenderer(Type modelType);
        IFormControlRenderer GetFormControlRenderer(string typeName, Type modelType);
    }
}