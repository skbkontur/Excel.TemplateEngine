using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal class ExcelDocumentNumberingFormats : IExcelDocumentNumberingFormats
    {
        public ExcelDocumentNumberingFormats(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            cache = new Dictionary<NumberingFormatCacheItem, uint>();
        }

        public uint AddFormat(ExcelCellNumberingFormat format)
        {
            if (format == null)
                return 0;
            var cacheItem = new NumberingFormatCacheItem(format);
            uint formatId;
            if (cache.TryGetValue(cacheItem, out formatId))
                return formatId;
            if (stylesheet.NumberingFormats == null)
            {
                var numberingFormats = new NumberingFormats {Count = new UInt32Value(0u)};
                stylesheet.InsertAt(numberingFormats, 0);
            }
            formatId = ++stylesheet.NumberingFormats.Count;
            stylesheet.NumberingFormats.AppendChild(cacheItem.ToNumberingFormat(formatId));
            cache.Add(cacheItem, formatId);
            return formatId;
        }

        private readonly Stylesheet stylesheet;

        private readonly Dictionary<NumberingFormatCacheItem, uint> cache;
    }
}