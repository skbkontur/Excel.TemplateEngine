using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public interface IParser
    {
        [NotNull]
        object Parse([NotNull] ITableParser tableParser, [NotNull] Type modelType, [NotNull] RenderingTemplate template, Action<string, string> addFieldMapping);
    }
}