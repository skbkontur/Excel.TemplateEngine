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
            underlined = style.Underlined;
        }

        public bool Equals(FontStyleCacheItem other)
        {
            if(ReferenceEquals(null, other)) return false;
            if(ReferenceEquals(this, other)) return true;
            return Equals(color, other.color) && size == other.size && underlined.Equals(other.underlined);
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
                var hashCode = (color != null ? color.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ size.GetHashCode();
                hashCode = (hashCode * 397) ^ underlined.GetHashCode();
                return hashCode;
            }
        }

        public Font ToFont()
        {
            return new Font
                {
                    Color = color == null ? null : color.ToColor(),
                    FontSize = size == null ? null : new FontSize {Val = size},
                    Underline = underlined ? new Underline() : null
                };
        }

        private readonly ColorCacheItem color;
        private readonly int? size;
        private readonly bool underlined;
    }
}