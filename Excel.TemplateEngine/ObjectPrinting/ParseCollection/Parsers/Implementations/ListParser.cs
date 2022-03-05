using System;
using System.Collections.Generic;
using System.Linq;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations
{
    internal class ListParser
    {
        /// <summary>
        /// </summary>
        /// <param name="tableReader"></param>
        /// <param name="modelType"></param>
        /// <param name="pathToList">With array access.</param>
        /// <param name="itemBound"></param>
        /// <param name="itemProps"></param>
        /// <param name="addItem"></param>
        /// <param name="logValueParse"></param>
        /// <returns></returns>
        [NotNull]
        public void Parse([NotNull] LazyTableReader tableReader,
                          [NotNull] Type modelType,
                          [NotNull] ExcelTemplatePath pathToList,
                          (int start, int end) itemBound,
                          [NotNull] ExcelTemplatePath[] itemProps,
                          [NotNull] Action<Dictionary<ExcelTemplatePath, object>> addItem,
                          [NotNull] Action<int, ExcelTemplatePath, string> logValueParse)
        {
            var fullPropsPaths = itemProps.Select(x => new ExcelTemplatePath(pathToList.RawPath + x.RawPath))
                                          .ToArray();

            var itemDict = itemProps.ToDictionary(x => x, _ => (object)null);
            var parsedCount = 0;
            foreach (var row in tableReader)
            {
                var rowReader = new LazyRowReader(row);
                var itemCells = rowReader.Where(x =>
                    {
                        var cellIndex = new ExcelCellIndex(x.CellReference);
                        return itemBound.start < cellIndex.ColumnIndex && cellIndex.ColumnIndex < itemBound.end;
                    }).ToArray();

                for (int i = 0; i < itemProps.Length; i++)
                {
                    var propType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, fullPropsPaths[i]);

                    CellTextParser.TryParse(itemCells[i].InnerText, propType, out var parsedValue);
                    itemDict[itemProps[i]] = parsedValue;
                }

                addItem(itemDict);

                logValueParse(parsedCount, pathToList, string.Join(", ", itemDict));
            }
        }

        private static object GetDefault(Type type)
        {
            if (type.IsValueType)
                return Activator.CreateInstance(type);
            return null;
        }
    }
}