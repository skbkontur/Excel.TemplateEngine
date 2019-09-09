using System;

using Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    public interface IFormValueParser
    {
        object ParseOrDefault([NotNull] ITableParser tableParser, [NotNull] string name, [NotNull] Type modelType);
    }
}