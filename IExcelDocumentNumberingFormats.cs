using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
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

        private class NumberingFormatCacheItem : IEquatable<NumberingFormatCacheItem>
        {
            public NumberingFormatCacheItem(ExcelCellNumberingFormat format)
            {
                Precision = format.Precision;
            }

            public bool Equals(NumberingFormatCacheItem other)
            {
                return Precision == other.Precision;
            }

            public override bool Equals(object obj)
            {
                if(ReferenceEquals(null, obj)) return false;
                if(ReferenceEquals(this, obj)) return true;
                if(obj.GetType() != GetType()) return false;
                return Equals((NumberingFormatCacheItem)obj);
            }

            public override int GetHashCode()
            {
                return Precision;
            }

            private int Precision { get; set; }
        }
    }
}