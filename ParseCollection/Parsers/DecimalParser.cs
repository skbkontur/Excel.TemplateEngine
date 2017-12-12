using System;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class DecimalParser : IAtomicValueParser
    {
        public DecimalParser(bool nullable = false)
        {
            this.nullable = nullable;
        }

        [CanBeNull]
        public object TryParse([NotNull] ITableParser tableParser, [NotNull] Type modelType)
        {
            // todo (mpivko, 15.12.2017):
            /*if (decimal.TryParse(value, numberStyles, russianCultureInfo, out var result) || decimal.TryParse(value, numberStyles, CultureInfo.InvariantCulture, out result))
                return decimalConverter.ToDecimal(decimalConverter.ToString(result)); // todo (mpivko, 15.12.2017): */


            if(nullable)
            {
                if (!typeof(decimal?).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected decimal?");
                if (tableParser.TryParseAtomicValue(out decimal? nullableResult))
                    return nullableResult;
            }
            else
            {
                if (!typeof(decimal).IsAssignableFrom(modelType))
                    throw new ArgumentException($"modelType is {modelType} but expected decimal");
                if (tableParser.TryParseAtomicValue(out decimal result))
                    return result;
            }

            return null;
        }

        private readonly bool nullable;
    }
}
 