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
        /// <param name="tableReader">TableReader of target document which will be parsed.</param>
        /// <param name="templateListCells">Template cells with list items descriptions.</param>
        /// <param name="filterTemplateCells">Determines whether it's needed to filter templateListCells or not.</param>
        /// <param name="logger"></param>
        public static IReadOnlyList<TItem> Parse<TItem>([NotNull] LazyTableReader tableReader,
                                                        [NotNull, ItemNotNull] IEnumerable<SimpleCell> templateListCells,
                                                        bool filterTemplateCells,
                                                        [NotNull] ILog logger)
        {
            var itemType = typeof(TItem);

            var template = filterTemplateCells ? FilterTemplateCells(templateListCells) : templateListCells;
            var listTemplate = template.Select(x => (CellPosition : x.CellPosition, ItemPropPath : ExcelTemplatePath.FromRawExpression(x.CellValue).SplitForEnumerableExpansion().relativePathToItem))
                                       .ToArray();

            var itemPropPaths = listTemplate.Select(x => x.ItemPropPath)
                                            .ToArray();
            var dictToObject = ObjectConversionGenerator.BuildDictToObject(itemPropPaths, itemType);

            var result = new List<TItem>();

            var firstListCellPosition = listTemplate.First().CellPosition;
            var row = tableReader.TryReadRow(firstListCellPosition.RowIndex);
            while (row != null)
            {
                var itemDict = itemPropPaths.ToDictionary(x => x, _ => (object)null);
                var rowIsEmpty = true;
                foreach (var prop in listTemplate)
                {
                    var cellPosition = new CellPosition(row.RowIndex, prop.CellPosition.ColumnIndex);
                    var cell = row.TryReadCell(cellPosition);
                    if (cell == null || string.IsNullOrEmpty(cell.CellValue))
                        continue;

                    rowIsEmpty = false;

                    var propType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(itemType, prop.ItemPropPath);
                    if (!TextValueParser.TryParse(cell.CellValue, propType, out var parsedValue))
                    {
                        logger.Warn("Failed to parse value {CellValue} from {CellReference} with type='{PropType}'", new {CellValue = cell.CellValue, CellReference = cell.CellPosition.CellReference, PropType = propType});
                        continue;
                    }

                    itemDict[prop.ItemPropPath] = parsedValue;
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

        /// <summary>
        ///     Returns cells of the first met row and which value describes items of the first met enumerable. Cells is ordered by ColumnIndex.
        /// </summary>
        /// <param name="templateListCells">Template cells that forms list description.</param>
        /// <returns></returns>
        public static IEnumerable<SimpleCell> FilterTemplateCells(IEnumerable<SimpleCell> templateListCells)
        {
            var cellsWithPaths = templateListCells.Where(x => TemplateDescriptionHelper.IsCorrectValueDescription(x.CellValue))
                                                  .Select(x => (cell : x, path : ExcelTemplatePath.FromRawExpression(x.CellValue)))
                                                  .Where(x => x.path.HasArrayAccess)
                                                  .OrderBy(x => x.cell.CellPosition.ColumnIndex)
                                                  .ToArray();

            var firstTemplateItem = cellsWithPaths[0];
            var firstEnumerablePath = firstTemplateItem.path
                                                       .SplitForEnumerableExpansion()
                                                       .pathToEnumerable;
            return cellsWithPaths.Where(x => x.cell.CellPosition.RowIndex == firstTemplateItem.cell.CellPosition.RowIndex &&
                                             x.path.RawPath.StartsWith(firstEnumerablePath.RawPath))
                                 .Select(x => x.cell);
        }
    }
}