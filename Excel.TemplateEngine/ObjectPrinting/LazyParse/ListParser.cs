using System.Collections.Generic;
using System.Linq;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse
{
    public static class ListParser
    {
        /// <summary>
        ///     Parse tableReader from its current position for List&lt;&gt; until it meets empty row.
        /// </summary>
        /// <param name="tableReader"></param>
        /// <param name="parentType">Base object type.</param>
        /// <param name="templateListCells">Template cells with list items descriptions.</param>
        /// <param name="logger"></param>
        public static IReadOnlyList<TItem> Parse<TItem>([NotNull] LazyTableReader tableReader,
                                                        [NotNull, ItemNotNull] SimpleCell[] templateListCells,
                                                        [NotNull] ILog logger)
        {
            var itemType = typeof(TItem);

            var listTemplate = templateListCells.Select(x => (cellPosition : x.CellPosition, itemPropPath : ExcelTemplatePath.FromRawExpression(x.CellValue).SplitForEnumerableExpansion().relativePathToItem))
                                                .OrderBy(x => x.cellPosition.ColumnIndex)
                                                .ToArray();

            var itemPropPaths = listTemplate.Select(x => x.itemPropPath)
                                            .ToArray();
            var dictToObject = ObjectConversionGenerator.BuildDictToObject(itemPropPaths, itemType);

            var result = new List<TItem>();

            var firstListCellPosition = listTemplate.First().cellPosition;
            var row = tableReader.TryReadRow(firstListCellPosition.RowIndex);
            while (row != null)
            {
                var itemDict = itemPropPaths.ToDictionary(x => x, _ => (object)null);
                var rowIsEmpty = true;
                foreach (var prop in listTemplate)
                {
                    var cellPosition = new CellPosition(row.RowIndex, prop.cellPosition.ColumnIndex);
                    var cell = row.TryReadCell(cellPosition);
                    if (cell == null || string.IsNullOrEmpty(cell.CellValue))
                        continue;

                    rowIsEmpty = false;

                    var propType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(itemType, prop.itemPropPath);
                    if (!TextValueParser.TryParse(cell.CellValue, propType, out var parsedValue))
                    {
                        logger.Warn("Failed to parse value {CellValue} from {CellReference} with type='{PropType}'", new {CellValue = cell.CellValue, CellReference = cell.CellPosition.CellReference, PropType = propType});
                        continue;
                    }

                    itemDict[prop.itemPropPath] = parsedValue;
                }

                if (rowIsEmpty)
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
    }
}