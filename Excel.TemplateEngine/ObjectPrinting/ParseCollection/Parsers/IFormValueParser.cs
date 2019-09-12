using System;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    internal interface IFormValueParser
    {
        object ParseOrDefault([NotNull] ITableParser tableParser, [NotNull] string name, [NotNull] Type modelType);
    }
}