using System;
using System.Collections;
using System.Linq;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;
using SKBKontur.Catalogue.Objects;
using SKBKontur.Catalogue.ServiceLib.Logging;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class ClassParser : IClassParser
    {
        public ClassParser(IParserCollection parserCollection)
        {
            this.parserCollection = parserCollection;
        }

        private const int maxIEnumerableLen = 200;

        [NotNull]
        public TModel Parse<TModel>([NotNull] ITableParser tableParser, [NotNull] RenderingTemplate template, Action<string, string> addFieldMapping)
            where TModel : new()
        {
            var model = new TModel();

            foreach(var row in template.Content.Cells)
            {
                foreach(var cell in row)
                {
                    tableParser.PushState(cell.CellPosition, new Styler(cell));

                    var expression = cell.StringValue;

                    if(TemplateDescriptionHelper.Instance.IsCorrectValueDescription(expression))
                    {
                        ParseValue(tableParser, addFieldMapping, model, cell, new ExcelTemplateExpression(expression));
                        continue;
                    }
                    if(TemplateDescriptionHelper.Instance.IsCorrectFormValueDescription(expression))
                    {
                        ParseFormValue(tableParser, addFieldMapping, model, cell, new ExcelTemplateExpression(expression));
                        continue;
                    }

                    tableParser.PopState();
                }
            }

            return model;
        }

        private void ParseValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ICell cell, ExcelTemplateExpression expression)
        {
            var childSetter = ObjectPropertySettersExtractor.ExtractChildObjectSetter(model, expression.ChildObjectPath);

            var childModelPath = expression.ChildObjectPath;
            var childModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), childModelPath);

            if(childModelPath.HasArrayAccess)
            {
                ParseEnumerableValue(tableParser, addFieldMapping, model, expression, childSetter, childModelType);
            }
            else
            {
                ParseSingleValue(tableParser, addFieldMapping, cell, childSetter, childModelPath, childModelType);
            }
        }

        private void ParseSingleValue(ITableParser tableParser, Action<string, string> addFieldMapping, ICell cell, Action<object> childSetter, ExcelTemplatePath childModelPath, Type childModelType)
        {
            var parser = parserCollection.GetAtomicValueParser(childModelType);
            if(!parser.TryParse(tableParser, childModelType, out var parsedObject))
            {
                Log.For(this).Error($"Failed to parse value '{cell.StringValue}' with childModelType='{childModelType}' via AtomicValueParser");
                return;
            }
            childSetter(parsedObject);
            addFieldMapping(childModelPath.RawPath, cell.CellPosition.CellReference);
        }

        private void ParseEnumerableValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ExcelTemplateExpression expression, Action<object> childSetter, Type childModelType)
        {
            var (rawPathToEnumerable, childPath) = expression.ChildObjectPath.SplitForEnumerableExpansion();

            var pathToEnumerable = ExcelTemplatePath.FromRawPath(rawPathToEnumerable);

            var cleanPathToEnumerable = pathToEnumerable.WithoutArrayAccess();

            var childEnumerableType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), cleanPathToEnumerable);
            if(!typeof(IList).IsAssignableFrom(childEnumerableType))
                throw new Exception($"Only ILists are supported as collections, but tried to use '{childEnumerableType}'. (path: {cleanPathToEnumerable.RawPath})");

            var childObject = ObjectPropertiesExtractor.ExtractChildObject(model, pathToEnumerable);
            if(childObject != null && !(childObject is IList))
                throw new InvalidProgramStateException("Failed to cast child to IList, although we checked that it should be IList");
            var childEnumerable = (IList)childObject;

            var parser = parserCollection.GetEnumerableParser(childEnumerableType);

            var limit = childEnumerable?.Count ?? -1;
            var parsedList = parser.Parse(tableParser, childModelType, limit, (name, value) => addFieldMapping($"{cleanPathToEnumerable.RawPath}{name}.{childPath}", value));

            var lastNotNull = parsedList.FindLastIndex(x => x != null);
            parsedList = parsedList.Take(lastNotNull + 1).ToList();

            if(parsedList.Count > maxIEnumerableLen)
                throw new EnumerableTooLongException(maxIEnumerableLen);

            childSetter(parsedList);
        }

        private void ParseFormValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ICell cell, ExcelTemplateExpression expression)
        {
            var childSetter = ObjectPropertySettersExtractor.ExtractChildObjectSetter(model, expression.ChildObjectPath);

            var childModelPath = expression.ChildObjectPath;
            var childModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), childModelPath);
            var childFormControlType = ExtractFormControlType(cell);
            var childFormControlName = ExtractFormControlName(cell);

            if(childModelPath.HasArrayAccess)
                throw new NotSupportedExcelSerializationException("Enumerables are not supported for form controls");

            var parser = parserCollection.GetFormValueParser(childFormControlType, childModelType);
            if(!parser.TryParse(tableParser, childFormControlName, childModelType, out var parsedObject))
                throw new FormControlParsingException(childFormControlName);

            childSetter(parsedObject);
            addFieldMapping(childModelPath.RawPath, childFormControlName);
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

        private readonly IParserCollection parserCollection;
    }
}