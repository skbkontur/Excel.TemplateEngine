using System;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    internal interface IAtomicValueParser
    {
        bool TryParse([NotNull] ITableParser tableParser, [NotNull] Type itemType, [CanBeNull] out object result);
    }
}