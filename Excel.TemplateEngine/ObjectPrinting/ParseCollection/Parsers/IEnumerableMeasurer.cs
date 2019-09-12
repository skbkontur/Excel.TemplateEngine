using System;
using System.Collections.Generic;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    internal interface IEnumerableMeasurer
    {
        int GetLength([NotNull] ITableParser tableParser, [NotNull] Type modelType, IEnumerable<ICell> primaryParts);
    }
}