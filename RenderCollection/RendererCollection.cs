using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection
{
    public class RendererCollection : IRendererCollection
    {
        public RendererCollection(ITemplateCollection templateCollection)
        {
            this.templateCollection = templateCollection;
        }

        public IRenderer GetRenderer(Type modelType)
        {
            if(modelType == typeof(string))
                return new StringRenderer();
            if(modelType == typeof(int))
                return new IntRender();
            if(modelType == typeof(decimal))
                return new DecimalRender();
            if(modelType == typeof(double))
                return new DoubleRender();
            if(TypeCheckingHelper.Instance.IsEnumerable(modelType))
                return new EnumerableRenderer(this);
            return new ClassRenderer(templateCollection, this);
        }

        private readonly ITemplateCollection templateCollection;
    }
}