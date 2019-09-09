using System;

using Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers
{
    public class CheckBoxValueParser : IFormValueParser
    {
        public bool TryParse(ITableParser tableParser, string name, Type modelType, out object result)
        {
            if (!tableParser.TryParseCheckBoxValue(name, out var parseResult))
            {
                result = null;
                return false;
            }
            result = parseResult;
            return true;
        }

        public object ParseOrDefault(ITableParser tableParser, string name, Type modelType)
        {
            if (!TryParse(tableParser, name, modelType, out var result))
                result = false;
            return result;
        }
    }
}