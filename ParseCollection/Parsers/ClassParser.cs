using System;
using System.Collections;
using System.Linq;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.Exceptions;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class ClassParser : IClassParser
    {
        private readonly IParserCollection parserCollection;
        private const int maxIEnumerableLen = 200;

        public ClassParser(IParserCollection parserCollection)
        {
            this.parserCollection = parserCollection;
        }

        [NotNull]
        public TModel Parse<TModel>([NotNull] ITableParser tableParser, [NotNull] RenderingTemplate template, Action<string, string> addFieldMapping)
            where TModel : new()
        {
            var model = new TModel();
            
            foreach (var row in template.Content.Cells)
            {
                foreach (var cell in row)
                {
                    tableParser.PushState(cell.CellPosition, new Styler(cell));

                    var expression = cell.StringValue;

                    if (TemplateDescriptionHelper.Instance.IsCorrectValueDescription(expression))
                    {
                        ParseValue(tableParser, addFieldMapping, model, cell, expression);
                        continue;
                    }
                    if (TemplateDescriptionHelper.Instance.IsCorrectFormValueDescription(expression))
                    {
                        ParseFormValue(tableParser, addFieldMapping, model, cell, expression);
                        continue;
                    }

                    tableParser.PopState();
                }
            }

            return model;
        }

        private void ParseValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ICell cell, string expression)
        {
            var childSetter = ObjectPropertiesExtractor.ExtractChildObjectSetter(model, expression);
            
            var childModelPath = ObjectPropertiesExtractor.ExtractChildObjectPath(expression);
            var childModelType = ObjectPropertiesExtractor.ExtractChildObjectType(model, expression);
            
            if (ObjectPropertiesExtractor.NeedEnumerableExpansion(expression))
            {
                ParseEnumerableValue(tableParser, addFieldMapping, model, expression, childSetter, childModelType);
            }
            else
            {
                ParseSingleValue(tableParser, addFieldMapping, cell, childSetter, childModelPath, childModelType);
            }
        }

        private void ParseSingleValue(ITableParser tableParser, Action<string, string> addFieldMapping, ICell cell, Action<object> childSetter, string childModelPath, Type childModelType)
        {
            var parser = parserCollection.GetAtomicValueParser(childModelType);
            if (!parser.TryParse(tableParser, childModelType, out var parsedObject))
                return; // todo (mpivko, 29.01.2018): 
            childSetter(parsedObject);
            addFieldMapping(childModelPath, cell.CellPosition.CellReference);
        }

        private void ParseEnumerableValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, string expression, Action<object> childSetter, Type childModelType)
        {
            var (pathToEnumerable, childPath) = ObjectPropertiesExtractor.SplitForEnumerableExpansion(expression);
            
            var cleanPathToEnumerable = pathToEnumerable.Replace("[]", "");
            
            var childEnumerableType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model, cleanPathToEnumerable);
            if (!typeof(IList).IsAssignableFrom(childEnumerableType))
                throw new Exception($"Only ILists are supported as collections, but tried to use '{childEnumerableType}'. (path: {cleanPathToEnumerable})");

            var childObject = ObjectPropertiesExtractor.Instance.ExtractChildObjectViaPath(model, cleanPathToEnumerable);
            if (childObject != null && !(childObject is IList))
                throw new InvalidProgramStateException("Failed to cast child to IList, although we checked that it should be IList");
            var childEnumerable = (IList)childObject;
            
            var parser = parserCollection.GetEnumerableParser(childEnumerableType);
            
            var limit = childEnumerable?.Count ?? maxIEnumerableLen + 1;
            var parsedList = parser.Parse(tableParser, childModelType, limit, (name, value) => addFieldMapping($"{cleanPathToEnumerable}{name}.{childPath}", value));

            var lastNotNull = parsedList.FindLastIndex(x => x != null);
            parsedList = parsedList.Take(lastNotNull + 1).ToList();

            if(parsedList.Count > maxIEnumerableLen)
                throw new EnumerableTooLongException(maxIEnumerableLen);
            
            childSetter(parsedList);
        }

        private void ParseFormValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ICell cell, string expression)
        {
            var childSetter = ObjectPropertiesExtractor.ExtractChildObjectSetter(model, expression);

            var childModelPath = ObjectPropertiesExtractor.ExtractChildObjectPath(expression);
            var cleanChildModelPath = ObjectPropertiesExtractor.ExtractCleanChildObjectPath(expression);
            var childModelType = ObjectPropertiesExtractor.ExtractChildObjectType(model, expression);
            var childFormControlType = ExtractFormControlType(cell);
            var childFormControlName = ExtractFormControlName(cell);

            if(ObjectPropertiesExtractor.NeedEnumerableExpansion(expression))
                throw new NotSupportedExcelSerializationException("Enumerables are not supported for form controls");

            var parser = parserCollection.GetFormValueParser(childFormControlType, childModelType);
            if(!parser.TryParse(tableParser, childFormControlName, childModelType, out var parsedObject))
                throw new FormControlParsingException(childFormControlName);

            childSetter(parsedObject);
            addFieldMapping(childModelPath, childFormControlName /*todo mpivko consider calculation of form control postion to return cell here instead of control name*/);
        }

        [NotNull]
        private static string ExtractFormControlName([NotNull] ICell cell)
        {
            return TemplateDescriptionHelper.Instance.GetFormControlNameFromValueDescription(cell.StringValue) ??
                   throw new InvalidExcelTemplateException($"Invalid xlsx template. '{cell.StringValue}' is not a valid form control description.");
        }

        [NotNull]
        private static string ExtractFormControlType([NotNull] ICell cell)
        {
            return TemplateDescriptionHelper.Instance.GetFormControlTypeFromValueDescription(cell.StringValue) ??
                   throw new InvalidExcelTemplateException($"Invalid xlsx template. '{cell.StringValue}' is not a valid form control description.");
        }
    }
}