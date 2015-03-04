using System.Collections.Generic;

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

        public string GetSharedString(uint index)
        {
            var a = sharedStringTable.ChildElements[(int)index];
            return a.FirstChild.InnerText;
        }

        public void Save()
        {
            if(sharedStringTable != null)
                sharedStringTable.Save();
        }

        private readonly IDictionary<SharedStringCacheItem, uint> cache;

        private readonly SharedStringTable sharedStringTable;
    }
}