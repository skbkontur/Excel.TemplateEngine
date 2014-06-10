using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.CacheItems;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    internal interface IExcelSharedStrings
    {
        uint AddSharedString(FormattedStringValue value);
        void Save();
    }

    internal class ExcelSharedStrings : IExcelSharedStrings
    {
        public ExcelSharedStrings(SharedStringTable sharedStringTable)
        {
            this.sharedStringTable = sharedStringTable;
            cache = new Dictionary<SharedStringCacheItem, uint>();
        }

        public uint AddSharedString(FormattedStringValue value)
        {
            var cacheItem = new SharedStringCacheItem(value);
            uint result;
            if(!cache.TryGetValue(cacheItem, out result))
            {
                result = sharedStringTable.UniqueCount;
                sharedStringTable.UniqueCount++;
                sharedStringTable.AppendChild(cacheItem.ToSharedStringItem());
                cache.Add(cacheItem, result);
            }
            return result;
        }

        public void Save()
        {
            sharedStringTable.Save();
        }

        private readonly IDictionary<SharedStringCacheItem, uint> cache;

        private readonly SharedStringTable sharedStringTable;
    }
}