using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations
{
    internal class ClassParser : IClassParser
    {
        public ClassParser(IParserCollection parserCollection, ILog logger)
        {
            this.parserCollection = parserCollection;
            this.logger = logger;
        }

        [NotNull]
        public TModel Parse<TModel>([NotNull] ITableParser tableParser, [NotNull] RenderingTemplate template, [NotNull] Action<string, string> addFieldMapping)
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
                    throw new InvalidOperationException($"Only ILists are supported as collections, but tried to use '{childEnumerableType}'. (path: {cleanPathToEnumerable.RawPath})");

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
            var modelType = model.GetType();
            var leafSetter = ObjectChildSetterFabric.GetChildObjectSetter(modelType, path);
            var leafModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, path);

            if (path.HasArrayAccess)
                ParseEnumerableValue(tableParser, addFieldMapping, model, path, leafSetter, leafModelType, enumerablesLengths);
            else
                ParseSingleValue(tableParser, addFieldMapping, model, leafSetter, path, leafModelType);
        }

        private void ParseSingleValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, Action<object, object> leafSetter, ExcelTemplatePath childModelPath, Type childModelType)
        {
            addFieldMapping(childModelPath.RawPath, tableParser.CurrentState.Cursor.CellReference);
            if (!TextValueParser.TryParse(tableParser.GetCurrentCellText(), childModelType, out var parsedObject))
            {
                logger.Error($"Failed to parse value from '{tableParser.CurrentState.Cursor.CellReference}' with childModelType='{childModelType}' via AtomicValueParser");
                return;
            }
            leafSetter(model, parsedObject);
        }

        private void ParseEnumerableValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ExcelTemplatePath path, Action<object, object> leafSetter, Type leafModelType, Dictionary<ExcelTemplatePath, int> enumerablesLengths)
        {
            var (rawPathToEnumerable, childPath) = path.SplitForEnumerableExpansion();

            var cleanPathToEnumerable = rawPathToEnumerable.WithoutArrayAccess();

            var enumerableType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), cleanPathToEnumerable);
            if (!typeof(IList).IsAssignableFrom(enumerableType))
                throw new Exception($"Only ILists are supported as collections, but tried to use '{enumerableType}'. (path: {cleanPathToEnumerable.RawPath})");

            var parser = parserCollection.GetEnumerableParser(enumerableType);

            var count = enumerablesLengths[cleanPathToEnumerable];
            var parsedList = parser.Parse(tableParser, leafModelType, count, (name, value) => addFieldMapping($"{cleanPathToEnumerable.RawPath}{name}.{childPath.RawPath}", value));

            leafSetter(model, parsedList);
        }

        private void ParseFormValue(ITableParser tableParser, Action<string, string> addFieldMapping, object model, ICell cell, ExcelTemplatePath path)
        {
            var modelType = model.GetType();
            var childSetter = ObjectChildSetterFabric.GetChildObjectSetter(modelType, path);
            var childModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, path);
            var (childFormControlType, childFormControlName) = GetFormControlDescription(cell);

            if (path.HasArrayAccess)
                throw new InvalidOperationException("Enumerables are not supported for form controls");

            var parser = parserCollection.GetFormValueParser(childFormControlType, childModelType);
            var parsedObject = parser.ParseOrDefault(tableParser, childFormControlName, childModelType);

            childSetter(model, parsedObject);
            addFieldMapping(path.RawPath, childFormControlName);
        }

        private static (string formControlType, string formControlName) GetFormControlDescription([NotNull] ICell cell)
        {
            var formControlDescription = TemplateDescriptionHelper.TryGetFormControlFromValueDescription(cell.StringValue);
            if (string.IsNullOrEmpty(formControlDescription.formControlType) || formControlDescription.formControlName == null)
                throw new InvalidOperationException($"Invalid xlsx template. '{cell.StringValue}' is not a valid form control description.");
            return formControlDescription;
        }

        private readonly IParserCollection parserCollection;
        private readonly ILog logger;
    }
}