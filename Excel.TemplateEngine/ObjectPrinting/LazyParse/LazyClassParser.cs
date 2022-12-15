using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
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
        /// <param name="readerOffset">Target file offset relative to a template.</param>
        [NotNull]
        public TModel Parse<TModel>([NotNull] LazyTableReader tableReader, [NotNull] RenderingTemplate template, ObjectSize readerOffset)
            where TModel : new()
        {
            var model = new TModel();

            foreach (var templateRow in template.Content.Cells)
            {
                var firstCell = templateRow.FirstOrDefault();
                if (firstCell == null)
                    continue;

                var targetRowReader = tableReader.TryReadRow(firstCell.CellPosition.RowIndex + readerOffset.Height);
                if (targetRowReader == null)
                    continue;

                using (targetRowReader)
                {
                    foreach (var templateCell in templateRow)
                    {
                        var targetCell = targetRowReader.TryReadCell(templateCell.CellPosition.Add(readerOffset));
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

                            var templateListCells = templateRow.SkipWhile(x => x.CellPosition.CellReference != templateCell.CellPosition.CellReference)
                                                               .ToArray();
                            ParseEnumerable(tableReader, model, templateListCells, enumerableType, readerOffset);
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
                                     [NotNull] [ItemNotNull] ICell[] templateListCells,
                                     [NotNull] Type enumerableType,
                                     [NotNull] ObjectSize readerOffset)
        {
            var firstEnumerablePath = ExcelTemplatePath.FromRawExpression(templateListCells.First().StringValue)
                                                       .SplitForEnumerableExpansion()
                                                       .pathToEnumerable;

            var modelType = model.GetType();
            var itemType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, firstEnumerablePath);

            var items = ParseList(tableReader, itemType, templateListCells.Select(x => new SimpleCell(x.CellPosition, x.StringValue)), readerOffset);

            var withoutArrayAccess = firstEnumerablePath.WithoutArrayAccess();
            var enumerableSetter = ObjectChildSetterFactory.GetEnumerableSetter(modelType, withoutArrayAccess, enumerableType, itemType);

            enumerableSetter(model, items);
        }

        private object ParseList([NotNull] LazyTableReader tableReader, [NotNull] Type itemType, [NotNull, ItemNotNull] IEnumerable<SimpleCell> templateListCells, ObjectSize readerOffset)
        {
            return parseList.MakeGenericMethod(itemType)
                            .Invoke(null, new object[] {tableReader, templateListCells, true, logger, readerOffset});
        }

        private void ParseSingleValue([NotNull] SimpleCell cell,
                                      [NotNull] object model,
                                      [NotNull] ExcelTemplatePath leafPath)
        {
            var leafSetter = ObjectChildSetterFactory.GetChildObjectSetter(model.GetType(), leafPath);
            var leafModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(model.GetType(), leafPath);

            if (!TextValueParser.TryParse(cell.CellValue, leafModelType, out var parsedObject))
            {
                logger.Warn("Failed to parse value {CellValue} from {CellReference} with type='{PropType}'", new {CellValue = cell.CellValue, CellReference = cell.CellPosition.CellReference, PropType = leafModelType});
                return;
            }

            leafSetter(model, parsedObject);
        }

        private readonly ILog logger;
        private readonly MethodInfo parseList = typeof(ListParser).GetMethod(nameof(ListParser.Parse));
    }
}