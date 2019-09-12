using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.Caches.CacheItems;
using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches.Implementations
{
    internal class ExcelDocumentBordersStyles : IExcelDocumentBordersStyles
    {
        public ExcelDocumentBordersStyles(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            cache = new Dictionary<BordersStyleCacheItem, uint>();
        }

        public uint AddStyle(ExcelCellBordersStyle format)
        {
            if (format == null)
                return 0;
            var cacheItem = new BordersStyleCacheItem(format);
            if (cache.TryGetValue(cacheItem, out var result))
                return result;
            if (stylesheet.Borders == null)
            {
                var borders = new Borders {Count = new UInt32Value(0u)};
                if (stylesheet.Fills != null)
                    stylesheet.InsertAfter(borders, stylesheet.Fills);
                else if (stylesheet.Fonts != null)
                    stylesheet.InsertAfter(borders, stylesheet.Fonts);
                else if (stylesheet.NumberingFormats != null)
                    stylesheet.InsertAfter(borders, stylesheet.NumberingFormats);
                else
                    stylesheet.InsertAt(borders, 0);
            }
            result = stylesheet.Borders.Count;
            stylesheet.Borders.AppendChild(cacheItem.ToBorder());
            stylesheet.Borders.Count++;
            cache.Add(cacheItem, result);
            return result;
        }

        private readonly Stylesheet stylesheet;

        private readonly Dictionary<BordersStyleCacheItem, uint> cache;
    }
}