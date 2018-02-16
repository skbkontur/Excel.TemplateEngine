using System.Linq;

using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public class ExcelTemplatePath
    {
        private ExcelTemplatePath(string rawPath)
        {
            if(!TemplateDescriptionHelper.Instance.IsCorrectModelPath(rawPath))
                throw new ObjectPropertyExtractionException($"Invalid excel template path '{rawPath}'");
            PartsWithIndexers = rawPath.Split('.');
            PartsWithoutArrayAccess = PartsWithIndexers.Select(x => x.Replace("[]", "")).ToArray();
            PartsWithoutIndexers = PartsWithIndexers.Select(x =>
                {
                    if(TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(x))
                        return TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartName(x);
                    return x.Replace("[]", "");
                }).ToArray();
            RawPath = rawPath;
            HasArrayAccess = PartsWithIndexers.Any(x => x.EndsWith("[]"));
        }

        public static ExcelTemplatePath FromRawPath(string rawPath)
        {
            return new ExcelTemplatePath(rawPath);
        }

        public ExcelTemplatePath WithoutArrayAccess()
        {
            return new ExcelTemplatePath(string.Join(".", PartsWithoutArrayAccess));
        }

        public bool HasArrayAccess { get; }
        public string RawPath { get; }
        public string[] PartsWithIndexers { get; }
        public string[] PartsWithoutArrayAccess { get; }
        public string[] PartsWithoutIndexers { get; }

        protected bool Equals(ExcelTemplatePath other)
        {
            return string.Equals(RawPath, other.RawPath);
        }

        public override bool Equals(object obj)
        {
            if(ReferenceEquals(null, obj)) return false;
            if(ReferenceEquals(this, obj)) return true;
            if(obj.GetType() != this.GetType()) return false;
            return Equals((ExcelTemplatePath)obj);
        }

        public override int GetHashCode()
        {
            return RawPath != null ? RawPath.GetHashCode() : 0;
        }
    }
}