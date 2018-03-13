using System.Collections.Concurrent;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal class ExcelDocumentFontStyles : IExcelDocumentFontStyles
    {
        public ExcelDocumentFontStyles(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            cache = new ConcurrentDictionary<FontStyleCacheItem, uint>();
        }

        public uint AddFont(ExcelCellFontStyle style)
        {
            if(style == null)
                return 0;
            var cacheItem = new FontStyleCacheItem(style);
            if(cache.TryGetValue(cacheItem, out var result))
                return result;
            if(stylesheet.Fonts == null)
            {
                var fonts = new Fonts {Count = new UInt32Value(0u)};
                if(stylesheet.NumberingFormats != null)
                    stylesheet.InsertAfter(fonts, stylesheet.NumberingFormats);
                else
                    stylesheet.InsertAt(fonts, 0);
            }
            result = stylesheet.Fonts.Count;
            stylesheet.Fonts.AppendChild(cacheItem.ToFont());
            stylesheet.Fonts.Count++;
            cache.TryAdd(cacheItem, result);
            return result;
        }

        private readonly Stylesheet stylesheet;
        private readonly ConcurrentDictionary<FontStyleCacheItem, uint> cache;
    }
}