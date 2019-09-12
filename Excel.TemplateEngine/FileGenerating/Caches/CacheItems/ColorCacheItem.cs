using System;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches.CacheItems
{
    internal class ColorCacheItem : IEquatable<ColorCacheItem>
    {
        public ColorCacheItem(ExcelColor color)
        {
            Red = color.Red;
            Green = color.Green;
            Blue = color.Blue;
            Alpha = color.Alpha;
        }

        public bool Equals(ColorCacheItem other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Red == other.Red && Green == other.Green && Blue == other.Blue && Alpha == other.Alpha;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((ColorCacheItem)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Red;
                hashCode = (hashCode * 397) ^ Green;
                hashCode = (hashCode * 397) ^ Blue;
                hashCode = (hashCode * 397) ^ Alpha;
                return hashCode;
            }
        }

        public T ToColor<T>() where T : ColorType, new()
        {
            return new T {Rgb = GetRGBString()};
        }

        private HexBinaryValue GetRGBString()
            => new HexBinaryValue {Value = $"{Alpha:X2}{Red:X2}{Green:X2}{Blue:X2}"};

        private int Red { get; }
        private int Green { get; }
        private int Blue { get; }
        private int Alpha { get; }
    }
}