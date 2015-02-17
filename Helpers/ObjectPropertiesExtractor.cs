using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public sealed class ObjectPropertiesExtractor
    {
        private ObjectPropertiesExtractor()
        {
        }

        public object ExtractChildObject(object model, string templateText)
        {
            if(!TemplateDescriptionHelper.Instance.IsCorrectValueDescription(templateText) ||
               model == null)
                return null;

            var descriptionParts = TemplateDescriptionHelper.Instance.GetDescriptionParts(templateText);
            var pathParts = descriptionParts[2].Split('.');

            return ExtractChildObject(model, pathParts);
        }

        public static ObjectPropertiesExtractor Instance { get { return instance; } }

        private static bool TryExtractCurrentChild(object model, string pathPart, out object child)
        {
            child = null;
            var propertyName = TemplateDescriptionHelper.Instance.GetPathPartName(pathPart);
            var propertyInfo = model.GetType()
                                    .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                                    .FirstOrDefault(property => property.Name == propertyName);
            if(propertyInfo == null)
                return false;

            child = propertyInfo.GetValue(model, null);
            return true;
        }

        private static void ExtractChildObject(object model, string[] pathParts, int pathPartIndex, List<object> result)
        {
            if(pathPartIndex == pathParts.Count())
            {
                result.Add(model);
                return;
            }

            object currentChild;
            if(!TryExtractCurrentChild(model, pathParts[pathPartIndex], out currentChild))
                return;

            if(currentChild == null)
            {
                result.Add(null);
                return;
            }

            if(TypeCheckingHelper.Instance.IsEnumerable(currentChild.GetType()))
            {
                foreach(var element in ((IEnumerable)currentChild).Cast<object>())
                    ExtractChildObject(element, pathParts, pathPartIndex + 1, result);
                return;
            }

            ExtractChildObject(currentChild, pathParts, pathPartIndex + 1, result);
        }

        private static object ExtractChildObject(object model, string[] pathParts)
        {
            var result = new List<object>();
            ExtractChildObject(model, pathParts, 0, result);
            if(pathParts.Any(TemplateDescriptionHelper.Instance.IsArrayPathPart))
                return result.ToArray();
            return result.Count == 0 ? null : result.Single();
        }

        private static readonly ObjectPropertiesExtractor instance = new ObjectPropertiesExtractor();
    }
}