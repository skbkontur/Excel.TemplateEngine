using System.Collections.Concurrent;

using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal class ExcelSharedStrings : IExcelSharedStrings
    {
        public ExcelSharedStrings(SharedStringTable sharedStringTable)
        {
            this.sharedStringTable = sharedStringTable;
            cache = new ConcurrentDictionary<SharedStringCacheItem, uint>();
        }

        public uint AddSharedString(FormattedStringValue value)
        {
            var cacheItem = new SharedStringCacheItem(value);
            if(!cache.TryGetValue(cacheItem, out var result))
            {
                if(sharedStringTable.UniqueCount == null)
                    sharedStringTable.UniqueCount = 0;
                result = sharedStringTable.UniqueCount;
                sharedStringTable.UniqueCount++;
                sharedStringTable.AppendChild(cacheItem.ToSharedStringItem());
                cache.TryAdd(cacheItem, result);
            }
            return result;
        }

        public string GetSharedString(uint index)
        {
            var a = sharedStringTable.ChildElements[(int)index];
            return a.FirstChild.InnerText;
        }

        private readonly ConcurrentDictionary<SharedStringCacheItem, uint> cache;

        private readonly SharedStringTable sharedStringTable;
    }
}