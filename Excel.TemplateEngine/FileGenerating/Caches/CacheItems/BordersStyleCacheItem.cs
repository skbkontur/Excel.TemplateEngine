using System;

using DocumentFormat.OpenXml.Spreadsheet;

using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Caches.CacheItems
{
    internal class BordersStyleCacheItem : IEquatable<BordersStyleCacheItem>
    {
        public BordersStyleCacheItem(ExcelCellBordersStyle format)
        {
            LeftBorder = format.LeftBorder == null ? null : new BorderCacheItem(format.LeftBorder);
            RightBorder = format.RightBorder == null ? null : new BorderCacheItem(format.RightBorder);
            TopBorder = format.TopBorder == null ? null : new BorderCacheItem(format.TopBorder);
            BottomBorder = format.BottomBorder == null ? null : new BorderCacheItem(format.BottomBorder);
        }

        public bool Equals(BordersStyleCacheItem other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(LeftBorder, other.LeftBorder) && Equals(RightBorder, other.RightBorder) && Equals(TopBorder, other.TopBorder) && Equals(BottomBorder, other.BottomBorder);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((BordersStyleCacheItem)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (LeftBorder != null ? LeftBorder.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (RightBorder != null ? RightBorder.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (TopBorder != null ? TopBorder.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (BottomBorder != null ? BottomBorder.GetHashCode() : 0);
                return hashCode;
            }
        }

        public Border ToBorder()
        {
            return new Border
                {
                    LeftBorder = LeftBorder?.ToBorder<LeftBorder>(),
                    RightBorder = RightBorder?.ToBorder<RightBorder>(),
                    TopBorder = TopBorder?.ToBorder<TopBorder>(),
                    BottomBorder = BottomBorder?.ToBorder<BottomBorder>()
                };
        }

        public BorderCacheItem LeftBorder { get; }
        public BorderCacheItem RightBorder { get; }
        public BorderCacheItem TopBorder { get; }
        public BorderCacheItem BottomBorder { get; }
    }
}