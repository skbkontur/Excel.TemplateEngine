using System;

using Excel.TemplateEngine.ObjectPrinting.TableParser;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    internal interface IAtomicValueParser
    {
        bool TryParse([NotNull] ITableParser tableParser, [NotNull] Type itemType, [CanBeNull] out object result);
    }
}