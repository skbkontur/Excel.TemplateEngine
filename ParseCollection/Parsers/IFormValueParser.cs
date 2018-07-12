using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public interface IFormValueParser
    {
        bool TryParse([NotNull] ITableParser tableParser, [NotNull] string name, [NotNull] Type modelType, out object result);
    }
}