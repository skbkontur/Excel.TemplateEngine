using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.DataTypes;
using Excel.TemplateEngine.FileGenerating.Implementation.CacheItems;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Caches
{
    internal class ExcelDocumentFillStyles : IExcelDocumentFillStyles
    {
        public ExcelDocumentFillStyles(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            cache = new Dictionary<FillStyleCacheItem, uint>();
        }

        public uint AddStyle(ExcelCellFillStyle format)
        {
            if (format == null)
                return 0;
            var cacheItem = new FillStyleCacheItem(format);
            uint result;
            if (cache.TryGetValue(cacheItem, out result))
                return result;
            if (stylesheet.Fills == null)
            {
                var fills = new Fills {Count = new UInt32Value(0u)};
                if (stylesheet.Fonts != null)
                    stylesheet.InsertAfter(fills, stylesheet.Fonts);
                else if (stylesheet.NumberingFormats != null)
                    stylesheet.InsertAfter(fills, stylesheet.NumberingFormats);
                else
                    stylesheet.InsertAt(fills, 0);
            }
            result = stylesheet.Fills.Count;
            stylesheet.Fills.AppendChild(cacheItem.ToFill());
            stylesheet.Fills.Count++;
            cache.Add(cacheItem, result);
            return result;
        }

        private readonly Stylesheet stylesheet;

        private readonly Dictionary<FillStyleCacheItem, uint> cache;
    }
}