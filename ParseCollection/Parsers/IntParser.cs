using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class IntParser : IAtomicValueParser
    {
        public IntParser(bool nullable=false)
        {
            this.nullable = nullable;
        }

        [CanBeNull]
        public object TryParse([NotNull] ITableParser tableParser, [NotNull] Type modelType)
        {
            if(nullable)
            {
                if (!typeof(int?).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected int?");
                if (tableParser.TryParseAtomicValue(out int? nullableResult))
                    return nullableResult;
            }
            else
            {
                if (!typeof(int).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected int");
                if (tableParser.TryParseAtomicValue(out int result))
                    return result;
            }

            return null;
        }

        private readonly bool nullable;
    }
}