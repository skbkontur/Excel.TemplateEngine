using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class CheckBoxValueParser : IFormValueParser
    {
        [CanBeNull]
        public object TryParse([NotNull] ITableParser tableParser, [NotNull] string name, [NotNull] Type modelType)
        {
            if(!tableParser.TryParseCheckBoxValue(name, out bool result))
                return null;
            return result;
        }
    }
}