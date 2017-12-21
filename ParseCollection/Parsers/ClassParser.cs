using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;
using SKBKontur.Catalogue.Expressions;
using SKBKontur.Catalogue.Linq;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class ClassParser : IClassParser
    {
        private readonly ITemplateCollection templateCollection;
        private readonly IParserCollection parserCollection;

        public ClassParser(ITemplateCollection templateCollection, IParserCollection parserCollection)
        {
            this.templateCollection = templateCollection;
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
            var cleanChildModelPath = ObjectPropertiesExtractor.ExtractCleanChildObjectPath(expression);
            var childModelType = ObjectPropertiesExtractor.ExtractChildObjectType(model, expression);
            var childTemplateName = ExtractTemplateName(cell);
            var maxIEnumerableLen = 200;

            if (ObjectPropertiesExtractor.NeedEnumerableExpansion(expression))
            {
                var (pathToEnumerable, childPath) = ObjectPropertiesExtractor.SplitForEnumerableExpansion(expression);
                var enumerableType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model, pathToEnumerable);

                var cleanPathToEnumerable = pathToEnumerable.Replace("[]", "");
                var childEnumerable = ObjectPropertiesExtractor.Instance.ExtractChildObjectViaPath(model, cleanPathToEnumerable);
                
                if (childEnumerable != null && !(childEnumerable is IList))
                    throw new Exception($"Trying to set IEnumerable to non-enumerable (name: {cleanPathToEnumerable}, type: {childEnumerable.GetType()})");

                if (!TypeCheckingHelper.Instance.IsEnumerable(enumerableType))
                    throw new Exception(); // todo (mpivko, 08.12.2017): ipse

                var parser = parserCollection.GetEnumerableParser(enumerableType);

                var clearPathToEnumerable = pathToEnumerable.Replace("[]", "");
                var limit = ((IList)childEnumerable)?.Count ?? maxIEnumerableLen + 1;
                var parsedObject = parser.Parse(tableParser, childModelType, limit, (name, value) => addFieldMapping($"{clearPathToEnumerable}{name}.{childPath}", value));
                var parsedList = (List<object>)parsedObject;

                var cntNullOrEmpty = 0;
                while (cntNullOrEmpty < parsedList.Count && parsedList[parsedList.Count - cntNullOrEmpty - 1] == null)
                    cntNullOrEmpty++;
                parsedList = parsedList.Take(parsedList.Count - cntNullOrEmpty).ToList();
                if (parsedList.Count > maxIEnumerableLen)
                {
                    // todo (mpivko, 18.12.2017): parsing error here, not exception
                    throw new Exception();
                }

                IList result;
                if (enumerableType.IsArray)
                {
                    var itemType = enumerableType.GetItemType();
                    result = Array.CreateInstance(childModelType, new[] { parsedList.Count });
                }
                else if (enumerableType.IsGenericType && enumerableType.GetGenericTypeDefinition() == typeof(List<>))
                {
                    var itemType = enumerableType.GetGenericArguments().Single();
                    result = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(itemType), new[] { parsedList.Count });
                }
                else
                {
                    throw new NotSupportedException();
                }

                for (int i = 0; i < parsedList.Count; i++)
                    result[i] = parsedList[i];
                // todo (mpivko, 11.12.2017): why dont pass parsedList there?
                childSetter(result);
            }
            else
            {
                try
                {
                    // todo (mpivko, 08.12.2017): copypaste
                    var parser = parserCollection.GetAtomicValueParser(childModelType);//TODO mpivko create ParserCaller
                    var parsedObject = parser.TryParse(tableParser, childModelType);
                    childSetter(parsedObject);
                    addFieldMapping(childModelPath, cell.CellPosition.CellReference);
                }
                catch (NotSupportedException)
                {
                    //TODO mpivko
                    throw;
                }
            }
        }

        private void ParseFormValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ICell cell, string expression)
        {
            var childSetter = ObjectPropertiesExtractor.ExtractChildObjectSetter(model, expression);

            var childModelPath = ObjectPropertiesExtractor.ExtractChildObjectPath(expression);
            var cleanChildModelPath = ObjectPropertiesExtractor.ExtractCleanChildObjectPath(expression);
            var childModelType = ObjectPropertiesExtractor.ExtractChildObjectType(model, expression);
            var childFormControlName = ExtractFormControlName(cell);

            if (ObjectPropertiesExtractor.NeedEnumerableExpansion(expression))
            {
                throw new NotSupportedException("Enumerables are not supported for form controls");
            }
            try
            {
                // todo (mpivko, 08.12.2017): copypaste
                var parser = parserCollection.GetFormValueParser(childModelType);//TODO mpivko create ParserCaller
                var parsedObject = parser.TryParse(tableParser, childFormControlName, childModelType);
                //todo mpivko parsedObject == null is error
                childSetter(parsedObject);
                addFieldMapping(childModelPath, childFormControlName /*todo mpivko consider calculation of form control postion to return cell here instead of control name*/);
            }
            catch (NotSupportedException)
            {
                //TODO mpivko
                throw;
            }
        }

        [NotNull]
        private static string ExtractTemplateName([NotNull] ICell cell)
        {
            return TemplateDescriptionHelper.Instance.ExtractTemplateNameFromValueDescription(cell.StringValue) ??
                   throw new InvalidProgramStateException($"Invalid xlsx template. '{cell.StringValue}' is not a valid value description.");
        }

        [NotNull]
        private static string ExtractFormControlName([NotNull] ICell cell)
        {
            return TemplateDescriptionHelper.Instance.ExtractFormControlNameFromValueDescription(cell.StringValue) ??
                   throw new InvalidProgramStateException($"Invalid xlsx template. '{cell.StringValue}' is not a valid form control description.");
        }

        private static void AssertCorrectValueDescription([CanBeNull] string valueDescription)
        {
            if (!TemplateDescriptionHelper.Instance.IsCorrectValueDescription(valueDescription))
                throw new InvalidProgramStateException($"Invalid xlsx template. '{valueDescription}' is not a valid value description.");
        }
    }
}