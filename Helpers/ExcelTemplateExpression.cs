using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public class ExcelTemplateExpression
    {
        public ExcelTemplateExpression([CanBeNull] string expression)
        {
            if(expression == null || !TemplateDescriptionHelper.Instance.IsCorrectAbstractValueDescription(expression))
                throw new ObjectPropertyExtractionException($"Invalid description '{expression}'");
            var parts = TemplateDescriptionHelper.Instance.GetDescriptionParts(expression);
            if(parts.Length != 3)
                throw new ObjectPropertyExtractionException($"Invalid description '{expression}'");
            ChildObjectPath = ExcelTemplatePath.FromRawPath(parts[2]);
        }

        public ExcelTemplatePath ChildObjectPath { get; }
    }
}