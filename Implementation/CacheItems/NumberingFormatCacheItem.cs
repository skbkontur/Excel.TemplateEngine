using System;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems
{
    public class NumberingFormatCacheItem : IEquatable<NumberingFormatCacheItem>
    {
        public NumberingFormatCacheItem(ExcelCellNumberingFormat format)
        {
            FormatCode = format.FormatCode;
        }

        public bool Equals(NumberingFormatCacheItem other)
        {
            return FormatCode == other.FormatCode;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((NumberingFormatCacheItem)obj);
        }

        public override int GetHashCode()
        {
            return FormatCode.GetHashCode();
        }

        public NumberingFormat ToNumberingFormat(uint formatId)
        {
            return new NumberingFormat
                {
                    FormatCode = new StringValue(FormatCode),
                    NumberFormatId = formatId
                };
        }

        public string FormatCode { get; private set; }
    }
}