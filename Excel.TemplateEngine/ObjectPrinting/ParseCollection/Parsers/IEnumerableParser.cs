using System;
using System.Collections.Generic;

using Excel.TemplateEngine.ObjectPrinting.TableParser;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    internal interface IEnumerableParser
    {
        [NotNull]
        List<object> Parse([NotNull] ITableParser tableParser, [NotNull] Type modelType, int count, [NotNull] Action<string, string> addFieldMapping);
    }
}