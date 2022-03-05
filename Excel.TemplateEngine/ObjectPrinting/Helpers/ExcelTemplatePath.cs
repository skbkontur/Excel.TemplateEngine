using System.Linq;

using SkbKontur.Excel.TemplateEngine.Exceptions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers
{
    internal class ExcelTemplatePath
    {
        public ExcelTemplatePath(string rawPath)
        {
            if (!TemplateDescriptionHelper.IsCorrectModelPath(rawPath))
                throw new ObjectPropertyExtractionException($"Invalid excel template path '{rawPath}'");
            PartsWithIndexers = rawPath.Split('.');
            PartsWithoutArrayAccess = PartsWithIndexers.Select(TemplateDescriptionHelper.GetArrayPathPartName).ToArray();
            RawPath = rawPath;
            HasArrayAccess = PartsWithIndexers.Any(TemplateDescriptionHelper.IsArrayPathPart);
            HasPrimaryKeyArrayAccess = PartsWithIndexers.Any(TemplateDescriptionHelper.IsPrimaryArrayPathPart);
        }

        public static ExcelTemplatePath FromRawExpression(string expression)
        {
            if (expression == null || !TemplateDescriptionHelper.IsCorrectAbstractValueDescription(expression))
                throw new ObjectPropertyExtractionException($"Invalid description '{expression}'");
            var parts = TemplateDescriptionHelper.GetDescriptionParts(expression);
            if (parts.Length != 3)
                throw new ObjectPropertyExtractionException($"Invalid description '{expression}'");
            return new ExcelTemplatePath(parts[2]);
        }

        public ExcelTemplatePath WithoutArrayAccess()
        {
            return new ExcelTemplatePath(string.Join(".", PartsWithoutArrayAccess));
        }

        public (ExcelTemplatePath pathToEnumerable, ExcelTemplatePath relativePathToItem) SplitForEnumerableExpansion()
        {
            if (!HasArrayAccess)
                throw new BaseExcelSerializationException($"Expression needs enumerable expansion but has no part with '[]' or '[#]' (path - '{RawPath}')");
            var pathToEnumerableLength = PartsWithIndexers.TakeWhile(x => !TemplateDescriptionHelper.IsArrayPathPart(x)).Count() + 1;
            var pathToEnumerable = new ExcelTemplatePath(string.Join(".", PartsWithIndexers.Take(pathToEnumerableLength)));
            var relativePathToItem = new ExcelTemplatePath(string.Join(".", PartsWithIndexers.Skip(pathToEnumerableLength)));
            return (pathToEnumerable, relativePathToItem);
        }

        public bool HasArrayAccess { get; }
        public bool HasPrimaryKeyArrayAccess { get; }
        public string RawPath { get; }
        public string[] PartsWithIndexers { get; }
        private string[] PartsWithoutArrayAccess { get; }

        private bool Equals(ExcelTemplatePath other)
        {
            return string.Equals(RawPath, other.RawPath);
        }

        public override bool Equals(object obj)
        {
            if (obj is null) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((ExcelTemplatePath)obj);
        }

        public override int GetHashCode()
        {
            return RawPath != null ? RawPath.GetHashCode() : 0;
        }
    }
}