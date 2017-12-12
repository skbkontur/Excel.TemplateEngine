using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class DropDownValueParser : IFormValueParser
    {
        public object TryParse(ITableParser tableParser, string name, Type modelType)
        {
            if (!tableParser.TryParseDropDownValue(name, out string result))
                return null;
            return result;
        }
    }
}