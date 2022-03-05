using System;
using System.Globalization;

using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations
{
    internal static class CellTextParser
    {
        public static bool TryParse([NotNull] string cellText, [NotNull] Type itemType, out object result)
        {
            if (itemType == typeof(string))
                return Parse(() => (TryParseString(cellText, out var res), res), out result);
            if (itemType == typeof(int))
                return Parse(() => (TryParseInt(cellText, out var res), res), out result);
            if (itemType == typeof(double))
                return Parse(() => (TryParseDouble(cellText, out var res), res), out result);
            if (itemType == typeof(decimal))
                return Parse(() => (TryParseDecimal(cellText, out var res), res), out result);
            if (itemType == typeof(long))
                return Parse(() => (TryParseLong(cellText, out var res), res), out result);
            if (itemType == typeof(int?))
                return Parse(() => (TryParseNullableInt(cellText, out var res), res), out result);
            if (itemType == typeof(double?))
                return Parse(() => (TryParseNullableDouble(cellText, out var res), res), out result);
            if (itemType == typeof(decimal?))
                return Parse(() => (TryParseNullableDecimal(cellText, out var res), res), out result);
            if (itemType == typeof(long?))
                return Parse(() => (TryParseNullableLong(cellText, out var res), res), out result);
            throw new InvalidOperationException($"Type {itemType} is not a supported atomic value");
        }

        private static bool TryParseString(string cellText, out string result)
        {
            result = cellText;
            return true;
        }

        private static bool TryParseInt(string cellText, out int result)
        {
            return int.TryParse(cellText, out result);
        }

        private static bool TryParseDouble(string cellText, out double result)
        {
            return double.TryParse(cellText, out result);
        }

        private static bool TryParseDecimal(string cellText, out decimal result)
        {
            return decimal.TryParse(cellText, numberStyles, russianCultureInfo, out result) ||
                   decimal.TryParse(cellText, numberStyles, CultureInfo.InvariantCulture, out result);
        }

        private static bool TryParseLong(string cellText, out long result)
        {
            return long.TryParse(cellText, out result);
        }

        private static bool TryParseNullableInt(string cellText, out int? result)
        {
            return TryParseNullableAtomicValue(cellText, x => (TryParseInt(x, out var res), res), out result);
        }

        private static bool TryParseNullableDouble(string cellText, out double? result)
        {
            return TryParseNullableAtomicValue(cellText, x => (TryParseDouble(x, out var res), res), out result);
        }

        private static bool TryParseNullableDecimal(string cellText, out decimal? result)
        {
            return TryParseNullableAtomicValue(cellText, x => (TryParseDecimal(x, out var res), res), out result);
        }

        private static bool TryParseNullableLong(string cellText, out long? result)
        {
            return TryParseNullableAtomicValue(cellText, x => (TryParseLong(x, out var res), res), out result);
        }

        private static bool TryParseNullableAtomicValue<T>(string cellText, Func<string, (bool success, T result)> parser, out T? result)
            where T : struct
        {
            if (string.IsNullOrEmpty(cellText))
            {
                result = null;
                return true;
            }

            bool success;
            (success, result) = parser(cellText);
            return success;
        }

        private static bool Parse<T>(Func<(bool success, T result)> parse, out object result)
        {
            var (success, res) = parse();
            result = res;
            return success;
        }

        private const NumberStyles numberStyles = NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign;

        private static readonly CultureInfo russianCultureInfo = CultureInfo.GetCultureInfo("ru-RU");
    }
}