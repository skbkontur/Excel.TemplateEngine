using System;
using System.Collections.Generic;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public interface IEnumerableMeasurer
    {
        int GetLength([NotNull] ITableParser tableParser, [NotNull] Type modelType, IEnumerable<ICell> primaryParts);
    }
}