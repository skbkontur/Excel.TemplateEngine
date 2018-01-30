using System.Linq;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public class ExpressionPath
    {
        private ExpressionPath(string rawPath)
        {
            PartsWithIndexers = rawPath.Split('.');
            PartsWithoutIndexers = PartsWithIndexers.Select(x =>
                {
                    if(TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(x))
                        return TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartName(x);
                    return x.Replace("[]", "");
                }).ToArray();
        }

        public static ExpressionPath FromRawPath(string rawPath)
        {
            return new ExpressionPath(rawPath);
        }

        public static ExpressionPath FromRawExpression(string rawExpression)
        {
            // todo (mpivko, 30.01.2018): 
            return FromRawPath(rawExpression.Split(':')[2]);
        }

        public string[] PartsWithIndexers { get; }
        public string[] PartsWithoutIndexers { get; }
    }
}