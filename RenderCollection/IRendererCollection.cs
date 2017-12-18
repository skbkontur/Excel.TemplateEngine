using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection
{
    public interface IRendererCollection
    {
        IRenderer GetRenderer(Type modelType);
        IFormControlRenderer GetFormControlRenderer(string typeName, Type modelType);
    }
}