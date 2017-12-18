using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class EnumerableParser : IEnumerableParser
    {
        public EnumerableParser(IParserCollection parserCollection)
        {
            this.parserCollection = parserCollection;
        }

        [NotNull]
        public IEnumerable Parse([NotNull] ITableParser tableParser, [NotNull] Type modelType, int count, [NotNull] Action<string, string> addFieldMapping)
        {
            if (count > maxEnumerableLength)
                throw new NotSupportedException($"Lists longer than {maxEnumerableLength} are not supported");
            
            var result = new List<object>();
            for (var i = 0; (count == -1 || i < count) && i < maxEnumerableLength; i++)
            {
                if (i != 0)
                    tableParser.MoveToNextLayer();

                tableParser.PushState();

                var parser = parserCollection.GetAtomicValueParser(modelType);
                var item = parser.TryParse(tableParser, modelType);
                if (count == -1 && item == null)
                    break;

                addFieldMapping($"[{i}]", tableParser.CurrentState.Cursor.CellReference);
                result.Add(item);
                tableParser.PopState();
            }
            
            return result;
        }

        private readonly IParserCollection parserCollection;
        private const int maxEnumerableLength = (int)1e4;
    }
}