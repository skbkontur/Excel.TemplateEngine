using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class AtomicValueParser : IAtomicValueParser
    {
        public bool TryParse([NotNull] ITableParser tableParser, [NotNull] Type itemType, out object result)
        {
            if (itemType == typeof(string))
                return Parse(() => (tableParser.TryParseAtomicValue(out string res), res), out result);
            if (itemType == typeof(int))
                return Parse(() => (tableParser.TryParseAtomicValue(out int res), res), out result);
            if (itemType == typeof(double))
                return Parse(() => (tableParser.TryParseAtomicValue(out double res), res), out result);
            if (itemType == typeof(decimal))
                return Parse(() => (tableParser.TryParseAtomicValue(out decimal res), res), out result);
            if (itemType == typeof(long))
                return Parse(() => (tableParser.TryParseAtomicValue(out long res), res), out result);
            if (itemType == typeof(int?))
                return Parse(() => (tableParser.TryParseAtomicValue(out int? res), res), out result);
            if (itemType == typeof(double?))
                return Parse(() => (tableParser.TryParseAtomicValue(out double? res), res), out result);
            if (itemType == typeof(decimal?))
                return Parse(() => (tableParser.TryParseAtomicValue(out decimal? res), res), out result);
            if (itemType == typeof(long?))
                return Parse(() => (tableParser.TryParseAtomicValue(out long? res), res), out result);
            throw new InvalidProgramStateException($"Type {itemType} is not a supported atomic value");
        }

        private bool Parse<T>(Func<(bool succeed, T result)> parse, out object result)
        {
            var (succeed, res) = parse();
            result = res;
            return succeed;
        }
    }
}