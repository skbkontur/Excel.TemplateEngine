using System;
using System.Linq;
using System.Text.RegularExpressions;

using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public sealed class TemplateDescriptionHelper
    {
        private TemplateDescriptionHelper()
        {
        }

        public string ExtractTemplateNameFromValueDescription(string expression)
        {
            return !IsCorrectValueDescription(expression) ? null : GetDescriptionParts(expression)[1];
        }

        public bool IsCorrectValueDescription(string expression)
        {
            var descriptionParts = GetDescriptionParts(expression);
            if(descriptionParts.Count() != 3 ||
               descriptionParts[0] != "Value" ||
               string.IsNullOrEmpty(descriptionParts[2]))
                return false;

            var pathRegex = new Regex(@"^[A-Za-z]\w*(\[\])?(\.[A-Za-z]\w*(\[\])?)*$");
            return pathRegex.IsMatch(descriptionParts[2]);
        }

        public bool IsCorrectTemplateDescription(string expression)
        {
            var descriptionParts = GetDescriptionParts(expression);
            if(descriptionParts.Count() != 4 ||
               descriptionParts[0] != "Template" ||
               string.IsNullOrEmpty(descriptionParts[1]))
                return false;

            var cellReferenceRegex = new Regex("^[A-Z]+[1-9][0-9]*$");
            return cellReferenceRegex.IsMatch(descriptionParts[2]) &&
                   cellReferenceRegex.IsMatch(descriptionParts[3]);
        }

        public bool TryExtractCoordinates(string templateDescription, out Tuple<ICellPosition, ICellPosition> range)
        {
            range = null;
            if(!IsCorrectTemplateDescription(templateDescription))
                return false;

            var cellReferenceRegex = new Regex("[A-Z]+[1-9][0-9]*");
            var upperLeft = new CellPosition(cellReferenceRegex.Matches(templateDescription)[0].Value);
            var lowerRight = new CellPosition(cellReferenceRegex.Matches(templateDescription)[1].Value);
            range = new Tuple<ICellPosition, ICellPosition>(upperLeft, lowerRight);

            return true;
        }

        public string[] GetDescriptionParts(string filedDescription)
        {
            return string.IsNullOrEmpty(filedDescription) ? new string[0] : filedDescription.Split(':').ToArray();
        }

        public string GetPathPartName(string pathPart)
        {
            return IsArrayPathPart(pathPart) ? pathPart.Replace("[]", "") : pathPart;
        }

        public bool IsArrayPathPart(string pathPart)
        {
            return pathPart.Contains("[]");
        }

        public static TemplateDescriptionHelper Instance { get { return instance; } }

        private static readonly TemplateDescriptionHelper instance = new TemplateDescriptionHelper();
    }
}