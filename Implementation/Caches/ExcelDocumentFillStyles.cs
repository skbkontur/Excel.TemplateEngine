using System.Collections.Concurrent;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal class ExcelDocumentFillStyles : IExcelDocumentFillStyles
    {
        public ExcelDocumentFillStyles(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            cache = new ConcurrentDictionary<FillStyleCacheItem, uint>();
        }

        public uint AddStyle(ExcelCellFillStyle format)
        {
            if(format == null)
                return 0;
            var cacheItem = new FillStyleCacheItem(format);
            if(cache.TryGetValue(cacheItem, out var result))
                return result;
            if(stylesheet.Fills == null)
            {
                var fills = new Fills {Count = new UInt32Value(0u)};
                if(stylesheet.Fonts != null)
                    stylesheet.InsertAfter(fills, stylesheet.Fonts);
                else if(stylesheet.NumberingFormats != null)
                    stylesheet.InsertAfter(fills, stylesheet.NumberingFormats);
                else
                    stylesheet.InsertAt(fills, 0);
            }
            result = stylesheet.Fills.Count;
            stylesheet.Fills.AppendChild(cacheItem.ToFill());
            stylesheet.Fills.Count++;
            cache.TryAdd(cacheItem, result);
            return result;
        }

        private readonly Stylesheet stylesheet;

        private readonly ConcurrentDictionary<FillStyleCacheItem, uint> cache;
    }
}