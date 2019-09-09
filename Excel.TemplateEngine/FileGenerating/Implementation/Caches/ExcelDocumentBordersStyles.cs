using System.Collections.Generic;

using Excel.TemplateEngine.FileGenerating.DataTypes;
using Excel.TemplateEngine.FileGenerating.Implementation.CacheItems;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Caches
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
            uint result;
            if (cache.TryGetValue(cacheItem, out result))
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