using System;
using System.Globalization;

using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers
{
    public static class TextValueParser
    {
        public static bool TryParse([CanBeNull] string cellText, [NotNull] Type itemType, out object result)
        {
            if (itemType == typeof(string))
                return TryParse(cellText, ParseString, out result);
            if (itemType == typeof(int))
                return TryParse(cellText, ParseInt, out result);
            if (itemType == typeof(double))
                return TryParse(cellText, ParseDouble, out result);
            if (itemType == typeof(decimal))
                return TryParse(cellText, ParseDecimal, out result);
            if (itemType == typeof(long))
                return TryParse(cellText, ParseLong, out result);

            if (itemType == typeof(int?))
                return TryParseNullable(cellText, ParseInt, out result);
            if (itemType == typeof(double?))
                return TryParseNullable(cellText, ParseDouble, out result);
            if (itemType == typeof(decimal?))
                return TryParseNullable(cellText, ParseDecimal, out result);
            if (itemType == typeof(long?))
                return TryParseNullable(cellText, ParseLong, out result);

            throw new InvalidOperationException($"Type {itemType} is not a supported atomic value");
        }

        private static (bool parsed, object res) ParseString(string cellText)
        {
            return (true, cellText);
        }

        private static (bool parsed, object res) ParseInt(string cellText)
        {
            return (int.TryParse(cellText, out var res), res);
        }

        private static (bool parsed, object res) ParseDouble(string cellText)
        {
            return (double.TryParse(cellText, out var res), res);
        }

        private static (bool parsed, object res) ParseDecimal(string cellText)
        {
            return (decimal.TryParse(cellText, numberStyles, russianCultureInfo, out var res) || decimal.TryParse(cellText, numberStyles, CultureInfo.InvariantCulture, out res), res);
        }

        private static (bool parsed, object res) ParseLong(string cellText)
        {
            return (long.TryParse(cellText, numberStyles, russianCultureInfo, out var res), res);
        }

        private static bool TryParse(string cellText, Func<string, (bool parsed, object res)> parse, out object result)
        {
            bool parsed;
            (parsed, result) = parse(cellText);
            return parsed;
        }

        private static bool TryParseNullable(string cellText, Func<string, (bool parsed, object res)> parse, out object result)
        {
            if (string.IsNullOrEmpty(cellText))
            {
                result = null;
                return true;
            }

            return TryParse(cellText, parse, out result);
        }

        private const NumberStyles numberStyles = NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign;

        private static readonly CultureInfo russianCultureInfo = CultureInfo.GetCultureInfo("ru-RU");
    }
}