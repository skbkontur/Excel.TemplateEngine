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
            PartsWithoutArrayAccess = PartsWithIndexers.Select(TemplateDescriptionHelper.Instance.GetArrayPathPartName).ToArray();
            PartsWithoutIndexers = PartsWithIndexers.Select(TemplateDescriptionHelper.Instance.GetPathPartName).ToArray();
            RawPath = rawPath;
            HasArrayAccess = PartsWithIndexers.Any(TemplateDescriptionHelper.Instance.IsArrayPathPart);
            HasPrimaryArrayAccess = PartsWithIndexers.Any(TemplateDescriptionHelper.Instance.IsPrimaryArrayPathPart);
        }

        public static ExcelTemplatePath FromRawPath(string rawPath)
        {
            return new ExcelTemplatePath(rawPath);
        }

        public ExcelTemplatePath WithoutArrayAccess()
        {
            return new ExcelTemplatePath(string.Join(".", PartsWithoutArrayAccess));
        }

        public (string, string) SplitForEnumerableExpansion()
        {
            if (!HasArrayAccess)
                throw new BaseExcelSerializationException($"Expression needs enumerable expansion but has no part with '[]' or '[#]' (path - '{RawPath}')");
            var parts = PartsWithIndexers;
            var firstPartLen = parts.TakeWhile(x => !TemplateDescriptionHelper.Instance.IsArrayPathPart(x)).Count() + 1;
            return (string.Join(".", parts.Take(firstPartLen)), string.Join(".", parts.Skip(firstPartLen)));
        }

        public bool HasArrayAccess { get; }
        public bool HasPrimaryArrayAccess { get; set; }
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