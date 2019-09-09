using System;
using System.Collections.Generic;

using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    public interface IEnumerableMeasurer
    {
        int GetLength([NotNull] ITableParser tableParser, [NotNull] Type modelType, IEnumerable<ICell> primaryParts);
    }
}