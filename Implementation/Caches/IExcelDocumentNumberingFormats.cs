using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    public interface IExcelDocumentNumberingFormats
    {
        uint AddFormat(ExcelCellNumberingFormat format);
    }

    internal class ExcelDocumentNumberingFormats : IExcelDocumentNumberingFormats
    {
        public ExcelDocumentNumberingFormats(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            cache = new Dictionary<NumberingFormatCacheItem, uint>();
        }

        public uint AddFormat(ExcelCellNumberingFormat format)
        {
            if(format == null)
                return 0;
            var cacheItem = new NumberingFormatCacheItem(format);
            uint result;
            if(cache.TryGetValue(cacheItem, out result))
                return result;
            var formatCode = "0." + string.Join("", Enumerable.Repeat("0", format.Precision));
            if(stylesheet.NumberingFormats == null)
            {
                var numberingFormats = new NumberingFormats {Count = new UInt32Value(0u)};
                stylesheet.InsertAt(numberingFormats, 0);
            }
            result = ++stylesheet.NumberingFormats.Count;
            stylesheet.NumberingFormats.AppendChild(new NumberingFormat
                {
                    FormatCode = new StringValue(formatCode),
                    NumberFormatId = result
                });
            return result;
        }

        private readonly Stylesheet stylesheet;

        private readonly Dictionary<NumberingFormatCacheItem, uint> cache;
    }
}