using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;

namespace SkbKontur.Excel.TemplateEngine.Extensions;

internal static class SortedDictionaryExtensions
{
    internal static IEnumerable<KeyValuePair<uint, Row>> RangeFromTo(
        this SortedDictionary<uint, Row> rowsCache,
        uint fromInclusive,
        uint toExclusive)
    {
        foreach (var kv in rowsCache) // неоптимально, т.к. ищет первое вхождение за O(n)
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
        if (rowsCache.TryGetValue(lookup, out var row))
        {
            successor = new KeyValuePair<uint, Row>(lookup, row);
            return true;
        }

        foreach (var kv in rowsCache) // тоже неоптимально, также ищет вхождение за O(n)
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