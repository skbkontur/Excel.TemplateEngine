using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public interface IAtomicValueParser // todo (mpivko, 15.12.2017): it's implementations just delegate calls to tableParser. Do we really need them?
    {
        [CanBeNull]
        object TryParse([NotNull] ITableParser tableParser, [NotNull] Type itemType);
    }
}