using System;
using System.Collections.Generic;
using System.Linq;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse
{
    internal static class ListParser
    {
        /// <summary>
        ///     Parse tableReader from its current position for List&lt;&gt; until it meets empty row.
        /// </summary>
        /// <param name="tableReader"></param>
        /// <param name="modelType">Base object type.</param>
        /// <param name="itemType"></param>
        /// <param name="templateListCells">Template cells with list items descriptions.</param>
        /// <param name="logger"></param>
        public static IReadOnlyList<object> Parse([NotNull] LazyTableReader tableReader,
                                                  [NotNull] Type modelType,
                                                  [NotNull] Type itemType,
                                                  [NotNull] [ItemNotNull] ICell[] templateListCells,
                                                  [NotNull] ILog logger)
        {
            var itemPropFullPaths = templateListCells.Select(x => (cellPosition : x.CellPosition, fullPropPath : ExcelTemplatePath.FromRawExpression(x.StringValue)))
                                                     .OrderBy(x => x.cellPosition.ColumnIndex)
                                                     .ToArray();
            var relativeItemPropsPaths = itemPropFullPaths.Select(x => x.fullPropPath.SplitForEnumerableExpansion().relativePathToItem)
                                                          .ToArray();
            var dictToObject = ObjectPropertySettersExtractor.GenerateDictToObjectFunc(relativeItemPropsPaths, itemType);

            var result = new List<object>();

            var firstListCellPosition = itemPropFullPaths.First().cellPosition;
            var row = tableReader.TryReadRow(firstListCellPosition.RowIndex);
            while (row != null)
            {
                var itemDict = relativeItemPropsPaths.ToDictionary(x => x, _ => (object)null);
                var rowIsEmpty = true;
                foreach (var prop in itemPropFullPaths)
                {
                    var cellPosition = new CellPosition(row.RowIndex, prop.cellPosition.ColumnIndex);
                    var cell = row.TryReadCell(cellPosition);
                    if (cell == null || string.IsNullOrEmpty(cell.CellValue))
                        continue;

                    rowIsEmpty = false;

                    var propType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, prop.fullPropPath);
                    if (!TextValueParser.TryParse(cell.CellValue, propType, out var parsedValue))
                    {
                        logger.Warn($"Failed to parse value {cell.CellValue} from {cell.CellPosition.CellReference} with type='{propType}'");
                        continue;
                    }

                    var relativeItemPropPath = prop.fullPropPath.SplitForEnumerableExpansion().relativePathToItem;
                    itemDict[relativeItemPropPath] = parsedValue;
                }

                if (rowIsEmpty)
                {
                    row.Dispose();
                    break;
                }

                result.Add(dictToObject(itemDict));

                row.Dispose();
                row = tableReader.TryReadRow(row.RowIndex + 1);
            }

            return result;
        }
    }
}