using System;

using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches.CacheItems
{
    internal class FontStyleCacheItem : IEquatable<FontStyleCacheItem>
    {
        public FontStyleCacheItem(ExcelCellFontStyle style)
        {
            size = style.Size;
            color = style.Color == null ? null : new ColorCacheItem(style.Color);
            underlined = style.Underlined;
            bold = style.Bold;
        }

        public bool Equals(FontStyleCacheItem other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(color, other.color) && size == other.size && underlined.Equals(other.underlined) && bold.Equals(other.bold);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((FontStyleCacheItem)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (color != null ? color.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ size.GetHashCode();
                hashCode = (hashCode * 397) ^ underlined.GetHashCode();
                hashCode = (hashCode * 397) ^ bold.GetHashCode();
                return hashCode;
            }
        }

        public Font ToFont()
        {
            return new Font
                {
                    Color = color?.ToColor<Color>(),
                    FontSize = size == null ? null : new FontSize {Val = size},
                    Underline = underlined ? new Underline() : null,
                    Bold = bold ? new Bold() : null
                };
        }

        private readonly ColorCacheItem color;
        private readonly int? size;
        private readonly bool underlined;
        private readonly bool bold;
    }
}