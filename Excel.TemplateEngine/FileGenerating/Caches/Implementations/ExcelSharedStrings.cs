using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;

using SkbKontur.Excel.TemplateEngine.FileGenerating.Caches.CacheItems;
using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Caches.Implementations
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
            if (!cache.TryGetValue(cacheItem, out var result))
            {
                if (sharedStringTable.UniqueCount == null)
                    sharedStringTable.UniqueCount = 0;
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
            return a.InnerText;
        }

        private readonly IDictionary<SharedStringCacheItem, uint> cache;

        private readonly SharedStringTable sharedStringTable;
    }
}