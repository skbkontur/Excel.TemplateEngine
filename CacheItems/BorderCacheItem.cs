using System;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator.CacheItems
{
    public class BorderCacheItem : IEquatable<BorderCacheItem>
    {
        public BorderCacheItem(ExcelCellBorderStyle borderStyle)
        {
            BorderType = borderStyle.BorderType;
            Color = new ColorCacheItem(borderStyle.Color);
        }

        public bool Equals(BorderCacheItem other)
        {
            if(ReferenceEquals(null, other)) return false;
            if(ReferenceEquals(this, other)) return true;
            return BorderType == other.BorderType && Equals(Color, other.Color);
        }

        public override bool Equals(object obj)
        {
            if(ReferenceEquals(null, obj)) return false;
            if(ReferenceEquals(this, obj)) return true;
            if(obj.GetType() != GetType()) return false;
            return Equals((BorderCacheItem)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((int)BorderType * 397) ^ (Color != null ? Color.GetHashCode() : 0);
            }
        }

        public Color GetColor()
        {
            return Color == null ? null : Color.ToColor();
        }

        public EnumValue<BorderStyleValues> GetStyle()
        {
            switch(BorderType)
            {
            case ExcelBorderType.None:
                return new EnumValue<BorderStyleValues>(BorderStyleValues.None);
            case ExcelBorderType.Single:
                return new EnumValue<BorderStyleValues>(BorderStyleValues.Medium);
            case ExcelBorderType.Bold:
                return new EnumValue<BorderStyleValues>(BorderStyleValues.Thick);
            case ExcelBorderType.Double:
                return new EnumValue<BorderStyleValues>(BorderStyleValues.Double);
            default:
                throw new Exception(string.Format("Unknown border type: {0}", BorderType));
            }
        }

        private ExcelBorderType BorderType { get; set; }
        private ColorCacheItem Color { get; set; }
    }
}