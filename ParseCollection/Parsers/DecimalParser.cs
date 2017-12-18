using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class DecimalParser : IAtomicValueParser
    {
        public DecimalParser(bool nullable = false)
        {
            this.nullable = nullable;
        }

        [CanBeNull]
        public object TryParse([NotNull] ITableParser tableParser, [NotNull] Type modelType)
        {
            if(nullable)
            {
                if (!typeof(decimal?).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected decimal?");
                if (tableParser.TryParseAtomicValue(out decimal? nullableResult))
                    return nullableResult;
            }
            else
            {
                if (!typeof(decimal).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected decimal");
                if (tableParser.TryParseAtomicValue(out decimal result))
                    return result;
            }

            return null;
        }

        private readonly bool nullable;
    }
}
 