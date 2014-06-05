using System;

namespace SKBKontur.Catalogue.ExcelFileGenerator.CacheItems
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

        private int Precision { get; set; }
    }
}