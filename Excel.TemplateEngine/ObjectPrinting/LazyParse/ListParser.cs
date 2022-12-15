using System;
using System.Collections.Generic;
using System.Linq;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse
{
    public static class ListParser
    {
        /// <summary>
        ///     Parse tableReader from its current position for List&lt;&gt; until it meets empty row.
        /// </summary>
        /// <param name="tableReader">TableReader of target document which will be parsed.</param>
        /// <param name="templateListCells">Template cells with list items descriptions.</param>
        /// <param name="filterTemplateCells">Determines whether it's needed to filter templateListCells or not.</param>
        /// <param name="readerOffset">Target file offset relative to a template.</param>
        /// <param name="logger"></param>
        public static IReadOnlyList<TItem> Parse<TItem>([NotNull] LazyTableReader tableReader,
                                                        [NotNull, ItemNotNull] IEnumerable<SimpleCell> templateListCells,
                                                        bool filterTemplateCells,
                                                        [NotNull] ILog logger,
                                                        [NotNull] ObjectSize readerOffset)
        {
            var itemType = typeof(TItem);

            var template = (filterTemplateCells ? FilterTemplateCells(templateListCells) : templateListCells).Select(x => (CellPosition : x.CellPosition, Path : ExcelTemplatePath.FromRawExpression(x.CellValue)))
                                                                                                             .ToArray();
            var impotentItemProps = template.Where(x => x.Path.HasPrimaryKeyArrayAccess)
                                            .Select(x => x.Path.SplitForEnumerableExpansion().relativePathToItem)
                                            .ToArray();
            var itemTemplate = template.Select(x => (CellPosition : x.CellPosition, ItemPropPath : x.Path.SplitForEnumerableExpansion().relativePathToItem))
                                       .ToArray();

            var itemPropPaths = itemTemplate.Select(x => x.ItemPropPath)
                                            .ToArray();
            var dictToObject = ObjectConversionGenerator.BuildDictToObject(itemPropPaths, itemType);

            var result = new List<TItem>();

            var firstItemCellPosition = itemTemplate.First().CellPosition;
            var row = tableReader.TryReadRow(firstItemCellPosition.RowIndex + readerOffset.Height);
            while (row != null)
            {
                var itemDict = itemPropPaths.ToDictionary(x => x, _ => (object)null);
                FillInItemDict(itemTemplate, row, itemType, itemDict, readerOffset, logger);

                if (IsRowEmpty(itemDict, impotentItemProps))
                {
                    row.Dispose();
                    break;
                }

                result.Add((TItem)dictToObject(itemDict));

                row.Dispose();
                row = tableReader.TryReadRow(row.RowIndex + 1);
            }

            return result;
        }

        private static void FillInItemDict((ICellPosition CellPosition, ExcelTemplatePath ItemPropPath)[] itemTemplate, 
                                           LazyRowReader row,
                                           Type itemType, 
                                           Dictionary<ExcelTemplatePath, object> itemDict, 
                                           ObjectSize readerOffset,
                                           ILog logger)
        {
            foreach (var prop in itemTemplate)
            {
                var cellPosition = new CellPosition(row.RowIndex, prop.CellPosition.ColumnIndex + readerOffset.Width);
                var cell = row.TryReadCell(cellPosition);
                if (cell == null || string.IsNullOrWhiteSpace(cell.CellValue))
                    continue;

                var propType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(itemType, prop.ItemPropPath);
                if (!TextValueParser.TryParse(cell.CellValue, propType, out var parsedValue))
                {
                    logger.Warn("Failed to parse value {CellValue} from {CellReference} with type='{PropType}'", new {CellValue = cell.CellValue, CellReference = cell.CellPosition.CellReference, PropType = propType});
                    continue;
                }

                itemDict[prop.ItemPropPath] = parsedValue;
            }
        }

        private static bool IsRowEmpty(Dictionary<ExcelTemplatePath, object> itemDict, ExcelTemplatePath[] impotentItemProps)
        {
            if (impotentItemProps.Length < 1)
                return itemDict.All(x => x.Value == null);

            return itemDict.Where(x => impotentItemProps.Contains(x.Key))
                           .All(x => x.Value == null);
        }

        /// <summary>
        ///     Returns cells of the first met row and which value describes items of the first met enumerable. Cells is ordered by ColumnIndex.
        /// </summary>
        /// <param name="templateListCells">Template cells that forms list description.</param>
        /// <returns></returns>
        [UsedImplicitly]
        public static IEnumerable<SimpleCell> FilterTemplateCells([NotNull, ItemNotNull] IEnumerable<SimpleCell> templateListCells)
        {
            var cellsWithPaths = templateListCells.Where(x => TemplateDescriptionHelper.IsCorrectValueDescription(x.CellValue))
                                                  .Select(x => (cell : x, path : ExcelTemplatePath.FromRawExpression(x.CellValue)))
                                                  .Where(x => x.path.HasArrayAccess)
                                                  .OrderBy(x => x.cell.CellPosition.ColumnIndex)
                                                  .ToArray();

            var firstTemplateItem = cellsWithPaths[0];
            var firstEnumerablePath = firstTemplateItem.path
                                                       .SplitForEnumerableExpansion()
                                                       .pathToEnumerable
                                                       .WithoutArrayAccess();
            return cellsWithPaths.Where(x => x.cell.CellPosition.RowIndex == firstTemplateItem.cell.CellPosition.RowIndex &&
                                             x.path.RawPath.StartsWith(firstEnumerablePath.RawPath))
                                 .Select(x => x.cell);
        }
    }
}