using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class LongParser : IAtomicValueParser
    {
        public LongParser(bool nullable = false)
        {
            this.nullable = nullable;
        }

        [CanBeNull]
        public object TryParse([NotNull] ITableParser tableParser, [NotNull] Type modelType)
        {
            if (nullable)
            {
                if (!typeof(long?).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected long?");
                if (tableParser.TryParseAtomicValue(out long? nullableResult))
                    return nullableResult;
            }
            else
            {
                if (!typeof(long).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected long");
                if (tableParser.TryParseAtomicValue(out long result))
                    return result;
            }

            return null;
        }

        private readonly bool nullable;
    }
}