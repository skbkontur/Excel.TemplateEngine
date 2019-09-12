using System;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches.CacheItems
{
    internal class FillStyleCacheItem : IEquatable<FillStyleCacheItem>
    {
        public FillStyleCacheItem(ExcelCellFillStyle format)
        {
            Color = new ColorCacheItem(format.Color ?? ExcelColors.White);
        }

        public bool Equals(FillStyleCacheItem other)
        {
            return Color.Equals(other.Color);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((FillStyleCacheItem)obj);
        }

        public override int GetHashCode() => Color.GetHashCode();

        public Fill ToFill()
            => new Fill
                {
                    PatternFill = new PatternFill(Color.ToColor<ForegroundColor>())
                        {
                            PatternType = new EnumValue<PatternValues>(PatternValues.Solid)
                        }
                };

        private ColorCacheItem Color { get; }
    }
}