using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class StringParser : IAtomicValueParser
    {
        [CanBeNull]
        public object TryParse([NotNull] ITableParser tableParser, [NotNull] Type modelType)
        {
            if (!typeof(string).IsAssignableFrom(modelType))
                throw new ArgumentException($"modelType is {modelType} but expected string");

            if(!tableParser.TryParseAtomicValue(out string result))
                return null;

            return result;
        }
    }
}