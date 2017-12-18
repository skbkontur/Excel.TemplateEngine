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
                return new IntRenderer();
            if(modelType == typeof(decimal))
                return new DecimalRenderer();
            if(modelType == typeof(double))
                return new DoubleRenderer();
            if(TypeCheckingHelper.Instance.IsEnumerable(modelType))
                return new EnumerableRenderer(this);
            return new ClassRenderer(templateCollection, this);
        }

        public IFormControlRenderer GetFormControlRenderer(string typeName, Type modelType)
        {
            // todo (mpivko, 19.12.2017): ;
            if(typeName == "CheckBox" && modelType == typeof(bool))
                return new CheckBoxRenderer();
            if(typeName == "DropDown" && modelType == typeof(string))
                return new DropDownRenderer();
            throw new Exception($"Unsupported pair of typeName ({typeName}) and modelType ({modelType}) for form controls");
        }

        private readonly ITemplateCollection templateCollection;
    }
}