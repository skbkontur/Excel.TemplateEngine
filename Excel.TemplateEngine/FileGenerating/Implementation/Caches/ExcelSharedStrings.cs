using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.DataTypes;
using Excel.TemplateEngine.FileGenerating.Implementation.CacheItems;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Caches
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
            return a.FirstChild.InnerText;
        }

        private readonly IDictionary<SharedStringCacheItem, uint> cache;

        private readonly SharedStringTable sharedStringTable;
    }
}