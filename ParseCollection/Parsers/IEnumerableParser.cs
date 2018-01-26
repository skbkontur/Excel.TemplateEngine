using System;
using System.Collections.Generic;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public interface IEnumerableParser
    {
        [NotNull]
        List<object> Parse([NotNull] ITableParser tableParser, [NotNull] Type modelType, int count, [NotNull] Action<string, string> addFieldMapping);
    }
}