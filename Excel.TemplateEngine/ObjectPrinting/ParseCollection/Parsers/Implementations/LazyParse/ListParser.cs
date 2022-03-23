using System;
using System.Collections.Generic;
using System.Linq;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations.LazyParse
{
    internal static class ListParser
    {
        public static void Parse([NotNull] LazyTableReader tableReader,
                                 [NotNull] Type modelType,
                                 [NotNull, ItemNotNull] ICell[] templateListCells,
                                 [NotNull] Action<Dictionary<ExcelTemplatePath, object>> addItem)
        {
            var parsedCount = 0;

            var itemPropFullPaths = templateListCells.Select(x => (cellPosition : x.CellPosition, fullPropPath : ExcelTemplatePath.FromRawExpression(x.StringValue)))
                                                     .OrderBy(x => x.cellPosition.ColumnIndex)
                                                     .ToArray();
            var firstListCellPosition = itemPropFullPaths.First().cellPosition;

            var row = tableReader.TryReadRow(firstListCellPosition.RowIndex);
            while (row != null)
            {
                var itemDict = itemPropFullPaths.ToDictionary(x => x.fullPropPath.SplitForEnumerableExpansion().relativePathToItem,
                                                              _ => (object)null);
                var rowIsEmpty = true;
                foreach (var prop in itemPropFullPaths)
                {
                    var cellPosition = new CellPosition(row.RowIndex, prop.cellPosition.ColumnIndex);
                    var cell = row.TryReadCell(cellPosition);
                    if (cell == null || string.IsNullOrEmpty(cell.CellValue))
                        continue;

                    rowIsEmpty = false;

                    var propType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, prop.fullPropPath);
                    CellTextParser.TryParse(cell.CellValue, propType, out var parsedValue);

                    var relativeItemPropPath = prop.fullPropPath.SplitForEnumerableExpansion().relativePathToItem;
                    itemDict[relativeItemPropPath] = parsedValue;
                }

                if (rowIsEmpty)
                    break;
                
                addItem(itemDict);

                row = tableReader.TryReadRow(row.RowIndex + 1);
            }
        }
    }
}