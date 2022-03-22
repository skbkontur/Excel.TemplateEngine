using System;
using System.Collections;
using System.Linq;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations.LazyParse
{
    internal class LazyClassParser
    {
        public LazyClassParser(ILog logger)
        {
            this.logger = logger;
        }

        [NotNull]
        public TModel Parse<TModel>([NotNull] LazyTableReader tableReader, [NotNull] RenderingTemplate template, [NotNull] Action<string, string> addFieldMapping)
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
                        var enumerableType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), path.WithoutArrayAccess());
                        if (!typeof(IList).IsAssignableFrom(enumerableType))
                            continue;

                        var templateListRow = templateRow.SkipWhile(x => x.CellPosition.CellReference != templateCell.CellPosition.CellReference)
                                                         .ToArray();
                        ParseList(tableReader, addFieldMapping, model, templateListRow);
                        break;
                    }

                    ParseSingleValue(targetCell, addFieldMapping, model, path);
                }
            }

            return model;
        }

        /// <summary>
        ///     Read rows one by one. Parse only first met enumerable elements.
        /// </summary>
        private void ParseList([NotNull] LazyTableReader tableReader,
                               [NotNull] Action<string, string> addFieldMapping,
                               [NotNull] object model,
                               [NotNull, ItemNotNull] ICell[] templateEnumerableRow)
        {
            var firstListExpression = templateEnumerableRow.First().StringValue;
            var listAccessEnd = firstListExpression.IndexOf(']');
            var listExpression = firstListExpression.Substring(0, listAccessEnd + 1);
            var pathToList = ExcelTemplatePath.FromRawExpression(listExpression);

            var firstListTemplateCells = templateEnumerableRow.Where(x => !string.IsNullOrEmpty(x.StringValue) &&
                                                                          x.StringValue.StartsWith(listExpression))
                                                              .ToArray();

            var relativeItemProps = firstListTemplateCells.Select(x => ExcelTemplatePath.FromRawExpression(x.StringValue)
                                                                                        .SplitForEnumerableExpansion()
                                                                                        .relativePathToItem)
                                                          .ToArray();
            var addItem = ObjectPropertySettersExtractor.GenerateChildListItemAdder(model, pathToList, relativeItemProps);

            var enumerableType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), pathToList.WithoutArrayAccess());

            ListParser.Parse(tableReader,
                             model.GetType(),
                             firstListTemplateCells,
                             addItem,
                             (index, value) => addFieldMapping($"{pathToList.WithoutArrayAccess()}[{index}]", value));
        }

        private void ParseSingleValue([NotNull] SimpleCell cell,
                                      [NotNull] Action<string, string> addFieldMapping,
                                      [NotNull] object model,
                                      [NotNull] ExcelTemplatePath leafPath)
        {
            var leafSetter = ObjectPropertySettersExtractor.ExtractChildObjectSetter(model, leafPath);
            var leafModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), leafPath);

            addFieldMapping(leafPath.RawPath, cell.CellPosition.CellReference);
            if (!CellTextParser.TryParse(cell.CellValue, leafModelType, out var parsedObject))
            {
                logger.Error($"Failed to parse value {cell.CellValue} from {cell.CellPosition.CellReference} with type='{leafModelType}'");
                return;
            }
            leafSetter(parsedObject);
        }

        private readonly ILog logger;
    }
}