using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class DoubleParser : IAtomicValueParser
    {
        public DoubleParser(bool nullable=false)
        {
            this.nullable = nullable;
        }

        [CanBeNull]
        public object TryParse([NotNull] ITableParser tableParser, [NotNull] Type modelType)
        {
            if (nullable)
            {
                if (!typeof(double?).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected double?");
                if (tableParser.TryParseAtomicValue(out double? nullableResult))
                    return nullableResult;
            }
            else
            {
                if (!typeof(double).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected double");
                if (tableParser.TryParseAtomicValue(out double result))
                    return result;
            }

            return null;
        }

        private readonly bool nullable;
    }
}