using System;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems
{
    public class NumberingFormatCacheItem : IEquatable<NumberingFormatCacheItem>
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

        public NumberingFormat ToNumberingFormat(uint formatId)
        {
            var formatCode = "0." + string.Join("", Enumerable.Repeat("0", Precision));
            return new NumberingFormat
                {
                    FormatCode = new StringValue(formatCode),
                    NumberFormatId = formatId
                };
        }

        private int Precision { get; set; }
    }
}