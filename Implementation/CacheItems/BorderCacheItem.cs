using System;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems
{
    public class BorderCacheItem : IEquatable<BorderCacheItem>
    {
        public BorderCacheItem(ExcelCellBorderStyle borderStyle)
        {
            borderType = borderStyle.BorderType;
            color = new ColorCacheItem(borderStyle.Color ?? ExcelColors.Black);
        }

        public bool Equals(BorderCacheItem other)
        {
            if(ReferenceEquals(null, other)) return false;
            if(ReferenceEquals(this, other)) return true;
            return borderType == other.borderType && Equals(color, other.color);
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
                return ((int)borderType * 397) ^ (color != null ? color.GetHashCode() : 0);
            }
        }

        public T ToBorder<T>() where T : BorderPropertiesType, new()
        {
            return new T
                {
                    Style = GetStyle(),
                    Color = GetColor()
                };
        }

        private Color GetColor()
        {
            return color?.ToColor<Color>();
        }

        private EnumValue<BorderStyleValues> GetStyle()
        {
            switch(borderType)
            {
            case ExcelBorderType.None:
                return new EnumValue<BorderStyleValues>(BorderStyleValues.None);
            case ExcelBorderType.Thin:
                return new EnumValue<BorderStyleValues>(BorderStyleValues.Thin);
            case ExcelBorderType.Single:
                return new EnumValue<BorderStyleValues>(BorderStyleValues.Medium);
            case ExcelBorderType.Bold:
                return new EnumValue<BorderStyleValues>(BorderStyleValues.Thick);
            case ExcelBorderType.Double:
                return new EnumValue<BorderStyleValues>(BorderStyleValues.Double);
            default:
                throw new Exception($"Unknown border type: {borderType}");
            }
        }

        private readonly ExcelBorderType borderType;
        private readonly ColorCacheItem color;
    }
}