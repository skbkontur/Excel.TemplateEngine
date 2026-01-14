using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;

namespace SkbKontur.Excel.TemplateEngine.Extensions
{
    internal static class SortedDictionaryExtensions
    {
        internal static IEnumerable<KeyValuePair<uint, Row>> RangeFromTo(
            this SortedDictionary<uint, Row> rowsCache,
            uint fromInclusive,
            uint toExclusive)
        {
            foreach (var kv in rowsCache)
            {
                if (kv.Key < fromInclusive) continue;
                if (kv.Key >= toExclusive) yield break;
                yield return kv;
            }
        }

        internal static bool TryWeakSuccessor(
            this SortedDictionary<uint, Row> rowsCache,
            uint lookup,
            out KeyValuePair<uint, Row> successor)
        {
            foreach (var kv in rowsCache)
            {
                if (kv.Key >= lookup)
                {
                    successor = kv;
                    return true;
                }
            }

            successor = default;
            return false;
        }
    }
}