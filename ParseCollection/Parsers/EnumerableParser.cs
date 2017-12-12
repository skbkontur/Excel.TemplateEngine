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
            /*if (!TypeCheckingHelper.Instance.IsEnumerable(modelType))
                throw new ArgumentException($"modelType is {modelType} but expected IEnumerable");

            var itemType = TypeCheckingHelper.Instance.GetEnumerableItemType(modelType);*/


            var maxIterations = (int)1e4; //TODO mpivko
            if (count > maxIterations)
                throw new NotSupportedException($"Lists longer than {maxIterations} are not supported");

            var result = new List<object>();
            for (var i = 0; i < count; i++)
            {
                if (i != 0)
                    tableParser.MoveToNextLayer();

                tableParser.PushState();

                Func<int, object> parse;
                try
                {
                    var parser = parserCollection.GetAtomicValueParser(modelType);
                    parse = _ => parser.TryParse(tableParser, modelType);
                }
                catch (NotSupportedException)
                {
                    throw new Exception($"There is no atomic value parser for '{modelType}'"); // todo (mpivko, 08.12.2017): 
                }
                
                var item = parse(i);

                /*if (isAtomicValue && item == null || !isAtomicValue && AllPropertiesAreNull(item))
                    break;*/
                addFieldMapping($"[{i}]", tableParser.CurrentState.Cursor.CellReference);
                result.Add(item);
                tableParser.PopState();
            }
            
            return result;
        }

        private bool AllPropertiesAreNull(object item)
        {
            return item.GetType().GetProperties().Select(x => x.GetValue(item)).All(x => x == null);
        }

        private readonly IParserCollection parserCollection;
    }
}