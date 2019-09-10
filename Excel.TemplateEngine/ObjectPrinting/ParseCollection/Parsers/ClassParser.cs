using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using Excel.TemplateEngine.Helpers;
using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using Excel.TemplateEngine.ObjectPrinting.TableParser;

using JetBrains.Annotations;

using Vostok.Logging.Abstractions;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    public class ClassParser : IClassParser
    {
        public ClassParser(IParserCollection parserCollection, ILog logger)
        {
            this.parserCollection = parserCollection;
            this.logger = logger;
        }

        [NotNull]
        public TModel Parse<TModel>([NotNull] ITableParser tableParser, [NotNull] RenderingTemplate template, Action<string, string> addFieldMapping)
            where TModel : new()
        {
            var model = new TModel();

            var enumerablesLengths = GetEnumerablesLengths<TModel>(tableParser, template);

            foreach (var row in template.Content.Cells)
            {
                foreach (var cell in row)
                {
                    tableParser.PushState(cell.CellPosition);

                    var expression = cell.StringValue;

                    if (TemplateDescriptionHelper.IsCorrectValueDescription(expression))
                    {
                        ParseCellularValue(tableParser, addFieldMapping, model, ExcelTemplatePath.FromRawExpression(expression), enumerablesLengths);
                        continue;
                    }
                    if (TemplateDescriptionHelper.IsCorrectFormValueDescription(expression))
                    {
                        ParseFormValue(tableParser, addFieldMapping, model, cell, ExcelTemplatePath.FromRawExpression(expression));
                        continue;
                    }

                    tableParser.PopState();
                }
            }

            return model;
        }

        [NotNull]
        private Dictionary<ExcelTemplatePath, int> GetEnumerablesLengths<TModel>([NotNull] ITableParser tableParser, [NotNull] RenderingTemplate template)
        {
            var enumerableCellsGroups = new Dictionary<ExcelTemplatePath, List<ICell>>();
            foreach (var row in template.Content.Cells)
            {
                foreach (var cell in row)
                {
                    var expression = cell.StringValue;

                    if (TemplateDescriptionHelper.IsCorrectValueDescription(expression) && ExcelTemplatePath.FromRawExpression(expression).HasArrayAccess)
                    {
                        var cleanPathToEnumerable = ExcelTemplatePath.FromRawExpression(expression)
                                                                     .SplitForEnumerableExpansion()
                                                                     .pathToEnumerable
                                                                     .WithoutArrayAccess();
                        if (!enumerableCellsGroups.ContainsKey(cleanPathToEnumerable))
                            enumerableCellsGroups[cleanPathToEnumerable] = new List<ICell>();
                        enumerableCellsGroups[cleanPathToEnumerable].Add(cell);
                    }
                }
            }

            var enumerablesLengths = new Dictionary<ExcelTemplatePath, int>();

            foreach (var enumerableCells in enumerableCellsGroups)
            {
                var cleanPathToEnumerable = enumerableCells.Key;

                var childEnumerableType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(typeof(TModel), cleanPathToEnumerable);
                if (!TypeCheckingHelper.IsIList(childEnumerableType))
                    throw new ExcelEngineException($"Only ILists are supported as collections, but tried to use '{childEnumerableType}'. (path: {cleanPathToEnumerable.RawPath})");

                var primaryParts = enumerableCells.Value.Where(x => ExcelTemplatePath.FromRawExpression(x.StringValue).HasPrimaryKeyArrayAccess).ToList();
                if (primaryParts.Count == 0)
                    primaryParts = enumerableCells.Value.Take(1).ToList();

                var measurer = parserCollection.GetEnumerableMeasurer();
                enumerablesLengths[cleanPathToEnumerable] = measurer.GetLength(tableParser, typeof(TModel), primaryParts);
            }

            return enumerablesLengths;
        }

        private void ParseCellularValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ExcelTemplatePath path, Dictionary<ExcelTemplatePath, int> enumerablesLengths)
        {
            var leafSetter = ObjectPropertySettersExtractor.ExtractChildObjectSetter(model, path);
            var leafModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), path);

            if (path.HasArrayAccess)
                ParseEnumerableValue(tableParser, addFieldMapping, model, path, leafSetter, leafModelType, enumerablesLengths);
            else
                ParseSingleValue(tableParser, addFieldMapping, leafSetter, path, leafModelType);
        }

        private void ParseSingleValue(ITableParser tableParser, Action<string, string> addFieldMapping, Action<object> leafSetter, ExcelTemplatePath childModelPath, Type childModelType)
        {
            var parser = parserCollection.GetAtomicValueParser();
            addFieldMapping(childModelPath.RawPath, tableParser.CurrentState.Cursor.CellReference);
            if (!parser.TryParse(tableParser, childModelType, out var parsedObject))
            {
                logger.Error($"Failed to parse value from '{tableParser.CurrentState.Cursor.CellReference}' with childModelType='{childModelType}' via AtomicValueParser");
                return;
            }
            leafSetter(parsedObject);
        }

        private void ParseEnumerableValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ExcelTemplatePath path, Action<object> leafSetter, Type leafModelType, Dictionary<ExcelTemplatePath, int> enumerablesLengths)
        {
            var (rawPathToEnumerable, childPath) = path.SplitForEnumerableExpansion();

            var cleanPathToEnumerable = rawPathToEnumerable.WithoutArrayAccess();

            var enumerableType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), cleanPathToEnumerable);
            if (!typeof(IList).IsAssignableFrom(enumerableType))
                throw new Exception($"Only ILists are supported as collections, but tried to use '{enumerableType}'. (path: {cleanPathToEnumerable.RawPath})");

            var parser = parserCollection.GetEnumerableParser(enumerableType);

            var count = enumerablesLengths[cleanPathToEnumerable];
            var parsedList = parser.Parse(tableParser, leafModelType, count, (name, value) => addFieldMapping($"{cleanPathToEnumerable.RawPath}{name}.{childPath.RawPath}", value));

            leafSetter(parsedList);
        }

        private void ParseFormValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ICell cell, ExcelTemplatePath path)
        {
            var childSetter = ObjectPropertySettersExtractor.ExtractChildObjectSetter(model, path);
            var childModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), path);
            var (childFormControlType, childFormControlName) = GetFormControlDescription(cell);

            if (path.HasArrayAccess)
                throw new ExcelEngineException("Enumerables are not supported for form controls");

            var parser = parserCollection.GetFormValueParser(childFormControlType, childModelType);
            var parsedObject = parser.ParseOrDefault(tableParser, childFormControlName, childModelType);

            childSetter(parsedObject);
            addFieldMapping(path.RawPath, childFormControlName);
        }

        private static (string formControlType, string formControlName) GetFormControlDescription([NotNull] ICell cell)
        {
            var formControlDescription = TemplateDescriptionHelper.TryGetFormControlFromValueDescription(cell.StringValue);
            if (string.IsNullOrEmpty(formControlDescription.formControlType) || formControlDescription.formControlName == null)
                throw new ExcelEngineException($"Invalid xlsx template. '{cell.StringValue}' is not a valid form control description.");
            return formControlDescription;
        }

        private readonly IParserCollection parserCollection;
        private readonly ILog logger;
    }
}