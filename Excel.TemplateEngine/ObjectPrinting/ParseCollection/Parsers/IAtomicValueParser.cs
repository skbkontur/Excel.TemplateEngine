using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public interface IAtomicValueParser
    {
        bool TryParse([NotNull] ITableParser tableParser, [NotNull] Type itemType, [CanBeNull] out object result);
    }
}