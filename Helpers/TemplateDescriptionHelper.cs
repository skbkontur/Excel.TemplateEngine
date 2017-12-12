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

        public string ExtractFormControlNameFromValueDescription(string expression)
        {
            return !IsCorrectFormValueDescription(expression) ? null : GetDescriptionParts(expression)[1];
        }

        public bool IsCorrectValueDescription(string expression)
        {
            var descriptionParts = GetDescriptionParts(expression);
            return IsCorrectAbstractValueDescription(expression) && descriptionParts[0] == "Value";
        }

        public bool IsCorrectFormValueDescription(string expression)
        {
            var formControlTypes = new[] {"CheckBox", "DropDown"}; // todo (mpivko, 15.12.2017): static hashset
            var descriptionParts = GetDescriptionParts(expression);
            return IsCorrectAbstractValueDescription(expression) && formControlTypes.Contains(descriptionParts[0]);
        }

        public bool IsCorrectAbstractValueDescription(string expression)
        {
            var descriptionParts = GetDescriptionParts(expression);
            if (descriptionParts.Count() != 3 ||
                string.IsNullOrEmpty(descriptionParts[2]))
                return false;

            var pathRegex = new Regex(@"^[A-Za-z]\w*(\[[^\[\]]*\])?(\.[A-Za-z]\w*(\[[^\[\]]*\])?)*$");
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

        public bool TryExtractCoordinates(string templateDescription, out IRectangle rectangle)
        {
            rectangle = null;
            if(!IsCorrectTemplateDescription(templateDescription))
                return false;

            rectangle = ExctractCoordinates(templateDescription);
            return true;
        }

        private static IRectangle ExctractCoordinates(string expression)
        {
            var cellReferenceRegex = new Regex("[A-Z]+[1-9][0-9]*");
            var upperLeft = new CellPosition(cellReferenceRegex.Matches(expression)[0].Value);
            var lowerRight = new CellPosition(cellReferenceRegex.Matches(expression)[1].Value);
            return new Rectangle(upperLeft, lowerRight);
        }

        public string[] GetDescriptionParts(string filedDescription)
        {
            return string.IsNullOrEmpty(filedDescription) ? new string[0] : filedDescription.Split(':').ToArray();
        }

        public string GetPathPartName(string pathPart)
        {
            if(IsArrayPathPart(pathPart))
                return pathPart.Replace("[]", "");
            if(IsCollectionAccessPathPart(pathPart))
                return GetCollectionAccessPathPartName(pathPart);
            return pathPart;
        }

        public bool IsArrayPathPart(string pathPart)
        {
            return pathPart.Contains("[]");
        }

        public static TemplateDescriptionHelper Instance { get { return instance; } }

        private static readonly TemplateDescriptionHelper instance = new TemplateDescriptionHelper();

        public string GetCollectionAccessPathPartName(string pathPart)
        {
            return GetCollectionAccessPathPart(pathPart).name;
        }

        public string GetCollectionAccessPathPartIndex(string pathPart)
        {
            return GetCollectionAccessPathPart(pathPart).index;
        }

        // todo (mpivko, 17.12.2017): copypaste
        private (string name, string index) GetCollectionAccessPathPart(string pathPart)
        {
            var regex = new Regex(@"^(\w*)\[([^\[\]]+)\]$"); // todo (mpivko, 17.12.2017): static and compiled
            var match = regex.Match(pathPart);
            if (!match.Success)
                throw new ArgumentException($"{nameof(pathPart)} should be collection access path part");
            return (match.Groups[1].Value, match.Groups[2].Value);
        }

        public bool IsCollectionAccessPathPart(string pathPart)
        {
            var regex = new Regex(@"^(\w*)\[([^\[\]]+)\]$");
            var match = regex.Match(pathPart);
            if (!match.Success)
                return false;
            return true;
        }
    }
}