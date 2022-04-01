using System;
using System.Linq;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse
{
    /// <summary>
    ///     Goes through first document sheet cell-by-cell from left to right from top to bottom without going back.
    ///     Can parse only separate cell values, List&lt;&gt; and array enumerations without size limitations.
    /// </summary>
    internal class LazyClassParser
    {
        public LazyClassParser(ILog logger)
        {
            this.logger = logger;
        }

        /// <summary>
        ///     Follows template and parse TModel from tableReader.
        /// </summary>
        /// <typeparam name="TModel">Class to parse.</typeparam>
        /// <param name="tableReader">Target document LazyTableReader.</param>
        /// <param name="template"></param>
        [NotNull]
        public TModel Parse<TModel>([NotNull] LazyTableReader tableReader, [NotNull] RenderingTemplate template)
            where TModel : new()
        {
            var model = new TModel();

            foreach (var templateRow in template.Content.Cells)
            {
                var firstCell = templateRow.FirstOrDefault();
                if (firstCell == null)
                    continue;

                var targetRowReader = tableReader.TryReadRow(firstCell.CellPosition.RowIndex);
                if (targetRowReader == null)
                    continue;

                using (targetRowReader)
                {
                    foreach (var templateCell in templateRow)
                    {
                        var targetCell = targetRowReader.TryReadCell(templateCell.CellPosition);
                        if (targetCell == null)
                            continue;

                        var expression = templateCell.StringValue;
                        if (!TemplateDescriptionHelper.IsCorrectValueDescription(expression))
                            continue;

                        var path = ExcelTemplatePath.FromRawExpression(expression);

                        if (path.HasArrayAccess)
                        {
                            var pathToEnumerable = path.SplitForEnumerableExpansion().pathToEnumerable;
                            var enumerableType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), pathToEnumerable.WithoutArrayAccess());
                            if (!(enumerableType.IsArray || TypeCheckingHelper.IsList(enumerableType)))
                                continue;

                            var templateEnumerableRow = templateRow.SkipWhile(x => x.CellPosition.CellReference != templateCell.CellPosition.CellReference)
                                                                   .ToArray();
                            ParseEnumerable(tableReader, model, templateEnumerableRow, enumerableType);
                            break;
                        }

                        ParseSingleValue(targetCell, model, path);
                    }
                }
            }

            return model;
        }

        /// <summary>
        ///     Read rows one-by-one. Parse only first met enumerable.
        /// </summary>
        private void ParseEnumerable([NotNull] LazyTableReader tableReader,
                                     [NotNull] object model,
                                     [NotNull] [ItemNotNull] ICell[] templateEnumerableRow,
                                     [NotNull] Type enumerableType)
        {
            var firstExpression = templateEnumerableRow.First().StringValue;
            var arrayAccessEnd = firstExpression.IndexOf(']');
            var enumerableExpression = firstExpression.Substring(0, arrayAccessEnd + 1);

            var firstEnumerableTemplateCells = templateEnumerableRow.Where(x => !string.IsNullOrEmpty(x.StringValue) &&
                                                                                x.StringValue.StartsWith(enumerableExpression))
                                                                    .ToArray();

            var modelType = model.GetType();
            var pathToEnumerable = ExcelTemplatePath.FromRawExpression(enumerableExpression);
            var itemType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, pathToEnumerable);

            var items = ListParser.Parse(tableReader, modelType, itemType, firstEnumerableTemplateCells, logger);

            var withoutArrayAccess = pathToEnumerable.WithoutArrayAccess();
            var enumerableSetter = ObjectChildSetterFabric.GetEnumerableSetter(modelType, withoutArrayAccess, enumerableType, itemType);

            enumerableSetter(model, items);
        }

        private void ParseSingleValue([NotNull] SimpleCell cell,
                                      [NotNull] object model,
                                      [NotNull] ExcelTemplatePath leafPath)
        {
            var leafSetter = ObjectChildSetterFabric.GetChildObjectSetter(model.GetType(), leafPath);
            var leafModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), leafPath);

            if (!TextValueParser.TryParse(cell.CellValue, leafModelType, out var parsedObject))
            {
                logger.Warn($"Failed to parse value {cell.CellValue} from {cell.CellPosition.CellReference} with type='{leafModelType}'");
                return;
            }

            leafSetter(model, parsedObject);
        }

        private readonly ILog logger;
    }
}