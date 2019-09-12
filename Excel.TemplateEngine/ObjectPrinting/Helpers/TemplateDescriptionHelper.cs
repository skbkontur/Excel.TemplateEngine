using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.Exceptions;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers
{
    internal static class TemplateDescriptionHelper
    {
        public static string GetTemplateNameFromValueDescription(string expression)
        {
            return !IsCorrectValueDescription(expression) ? null : GetDescriptionParts(expression)[1];
        }

        public static (string formControlType, string formControlName) TryGetFormControlFromValueDescription(string expression)
        {
            var descriptionParts = GetDescriptionParts(expression);
            return !IsCorrectFormValueDescription(expression) ? (null, null) : (descriptionParts[0], descriptionParts[1]);
        }

        public static bool IsCorrectValueDescription(string expression)
        {
            var descriptionParts = GetDescriptionParts(expression);
            return IsCorrectAbstractValueDescription(expression) && descriptionParts[0] == "Value";
        }

        public static bool IsCorrectFormValueDescription(string expression)
        {
            var descriptionParts = GetDescriptionParts(expression);
            return IsCorrectAbstractValueDescription(expression) && !string.IsNullOrEmpty(descriptionParts[1]) && formControlTypes.Contains(descriptionParts[0]);
        }

        public static bool IsCorrectAbstractValueDescription(string expression)
        {
            var descriptionParts = GetDescriptionParts(expression);
            if (descriptionParts.Count() != 3 ||
                string.IsNullOrEmpty(descriptionParts[2]))
                return false;

            return IsCorrectModelPath(descriptionParts[2]);
        }

        public static bool IsCorrectModelPath(string pathParts)
        {
            return pathRegex.IsMatch(pathParts);
        }

        public static bool IsCorrectTemplateDescription(string expression)
        {
            var descriptionParts = GetDescriptionParts(expression);
            if (descriptionParts.Count() != 4 ||
                descriptionParts[0] != "Template" ||
                string.IsNullOrEmpty(descriptionParts[1]))
                return false;

            return exactCellReferenceRegex.IsMatch(descriptionParts[2]) &&
                   exactCellReferenceRegex.IsMatch(descriptionParts[3]);
        }

        public static bool TryExtractCoordinates(string templateDescription, out IRectangle rectangle)
        {
            rectangle = null;
            if (!IsCorrectTemplateDescription(templateDescription))
                return false;

            rectangle = ExtractCoordinates(templateDescription);
            return true;
        }

        private static IRectangle ExtractCoordinates(string expression)
        {
            var upperLeft = new CellPosition(cellReferenceRegex.Matches(expression)[0].Value);
            var lowerRight = new CellPosition(cellReferenceRegex.Matches(expression)[1].Value);
            return new Rectangle(upperLeft, lowerRight);
        }

        public static string[] GetDescriptionParts(string filedDescription)
        {
            return string.IsNullOrEmpty(filedDescription) ? new string[0] : filedDescription.Split(':').ToArray();
        }

        public static string GetPathPartName(string pathPart)
        {
            if (IsArrayPathPart(pathPart))
                return GetArrayPathPartName(pathPart);
            if (IsCollectionAccessPathPart(pathPart))
                return GetCollectionAccessPathPartName(pathPart);
            return pathPart;
        }

        public static string GetArrayPathPartName(string pathPart)
        {
            return pathPart.Replace("[]", "").Replace("[#]", "");
        }

        public static bool IsArrayPathPart(string pathPart)
        {
            return arrayPathPartRegex.IsMatch(pathPart);
        }

        public static bool IsPrimaryArrayPathPart(string pathPart)
        {
            return arrayPrimaryPathPartRegex.IsMatch(pathPart);
        }

        public static string GetCollectionAccessPathPartName(string pathPart)
        {
            return GetCollectionAccessPathPart(pathPart).name;
        }

        public static string GetCollectionAccessPathPartIndex(string pathPart)
        {
            return GetCollectionAccessPathPart(pathPart).index;
        }

        public static bool IsCollectionAccessPathPart(string pathPart)
        {
            return collectionAccessPathPartRegex.IsMatch(pathPart);
        }

        private static (string name, string index) GetCollectionAccessPathPart(string pathPart)
        {
            var match = collectionAccessPathPartRegex.Match(pathPart);
            if (!match.Success)
                throw new ArgumentException($"{nameof(pathPart)} should be collection access path part");
            return (match.Groups[1].Value, match.Groups[2].Value);
        }

        [NotNull]
        public static object ParseCollectionIndexerOrThrow([NotNull] string collectionIndexer, [NotNull] Type collectionKeyType)
        {
            if (collectionIndexer.StartsWith("\"") && collectionIndexer.EndsWith("\""))
            {
                if (collectionKeyType != typeof(string))
                    throw new ObjectPropertyExtractionException($"Collection with '{collectionKeyType}' keys was indexed by {typeof(string)}");
                return collectionIndexer.Substring(1, collectionIndexer.Length - 2);
            }
            if (int.TryParse(collectionIndexer, out var intIndexer))
            {
                if (collectionKeyType != typeof(int))
                    throw new ObjectPropertyExtractionException($"Collection with '{collectionKeyType}' keys was indexed by {typeof(int)}");
                return intIndexer;
            }
            throw new ObjectPropertyExtractionException("Only strings and ints are supported as collection indexers");
        }

        private static readonly Regex collectionAccessPathPartRegex = new Regex(@"^(\w+)\[([^\[\]#]+)\]$", RegexOptions.Compiled);
        private static readonly Regex arrayPathPartRegex = new Regex(@"^(\w+)\[#?\]$", RegexOptions.Compiled);
        private static readonly Regex arrayPrimaryPathPartRegex = new Regex(@"^(\w+)\[#\]$", RegexOptions.Compiled);
        private static readonly Regex pathRegex = new Regex(@"^[A-Za-z]\w*(\[[^\[\]]*\])?(\.[A-Za-z]\w*(\[[^\[\]]*\])?)*$", RegexOptions.Compiled);
        private static readonly Regex cellReferenceRegex = new Regex("[A-Z]+[1-9][0-9]*", RegexOptions.Compiled);
        private static readonly Regex exactCellReferenceRegex = new Regex("^[A-Z]+[1-9][0-9]*$", RegexOptions.Compiled);
        private static readonly HashSet<string> formControlTypes = new HashSet<string>(new[] {"CheckBox", "DropDown"});
    }
}