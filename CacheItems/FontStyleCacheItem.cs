using System;

using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator.CacheItems
{
    public class FontStyleCacheItem : IEquatable<FontStyleCacheItem>
    {
        public FontStyleCacheItem(ExcelCellFontStyle style)
        {
            size = style.Size;
            color = style.Color == null ? null : new ColorCacheItem(style.Color);
        }

        public bool Equals(FontStyleCacheItem other)
        {
            if(ReferenceEquals(null, other)) return false;
            if(ReferenceEquals(this, other)) return true;
            return Equals(color, other.color) && Equals(size, other.size);
        }

        public Font ToFont()
        {
            return new Font
                {
                    Color = color == null ? null : color.ToColor(),
                    FontSize = size == null ? null : new FontSize {Val = size}
                };
        }

        public override bool Equals(object obj)
        {
            if(ReferenceEquals(null, obj)) return false;
            if(ReferenceEquals(this, obj)) return true;
            if(obj.GetType() != GetType()) return false;
            return Equals((FontStyleCacheItem)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((color != null ? color.GetHashCode() : 0) * 397) ^ (size ?? 0);
            }
        }

        private readonly ColorCacheItem color;
        private readonly int? size;
    }
}