using System;
using System.Collections.Generic;

using Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using Excel.TemplateEngine.ObjectPrinting.TableParser;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    internal interface IEnumerableMeasurer
    {
        int GetLength([NotNull] ITableParser tableParser, [NotNull] Type modelType, IEnumerable<ICell> primaryParts);
    }
}