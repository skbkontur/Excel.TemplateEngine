using System;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection
{
    internal class RendererCollection : IRendererCollection
    {
        public RendererCollection(ITemplateCollection templateCollection)
        {
            this.templateCollection = templateCollection;
        }

        [NotNull]
        public IRenderer GetRenderer([NotNull] Type modelType)
        {
            if (modelType == typeof(string))
                return new StringRenderer();
            if (modelType == typeof(int))
                return new IntRenderer();
            if (modelType == typeof(decimal))
                return new DecimalRenderer();
            if (modelType == typeof(double))
                return new DoubleRenderer();
            if (TypeCheckingHelper.IsEnumerable(modelType))
                return new EnumerableRenderer(this);
            return new ClassRenderer(templateCollection, this);
        }

        [NotNull]
        public IFormControlRenderer GetFormControlRenderer([NotNull] string typeName, [NotNull] Type modelType)
        {
            if (typeName == "CheckBox" && modelType == typeof(bool))
                return new CheckBoxRenderer();
            if (typeName == "DropDown" && modelType == typeof(string))
                return new DropDownRenderer();
            throw new InvalidOperationException($"Unsupported pair of typeName ({typeName}) and modelType ({modelType}) for form controls");
        }

        private readonly ITemplateCollection templateCollection;
    }
}