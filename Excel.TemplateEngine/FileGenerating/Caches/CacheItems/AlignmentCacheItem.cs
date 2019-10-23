using System;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Caches.CacheItems
{
    internal class AlignmentCacheItem : IEquatable<AlignmentCacheItem>
    {
        public AlignmentCacheItem(ExcelCellAlignment cellAlignment)
        {
            verticalAlignment = cellAlignment.VerticalAlignment;
            horizontalAlignment = cellAlignment.HorizontalAlignment;
            wrapText = cellAlignment.WrapText;
        }

        public bool Equals(AlignmentCacheItem other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return verticalAlignment == other.verticalAlignment && horizontalAlignment == other.horizontalAlignment && wrapText.Equals(other.wrapText);
        }

        public Alignment ToAlignment()
        {
            return new Alignment
                {
                    Horizontal = GetHorizontalAlignment(),
                    Vertical = GetVerticalAlignment(),
                    WrapText = wrapText ? new BooleanValue(true) : null
                };
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((AlignmentCacheItem)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (int)verticalAlignment;
                hashCode = (hashCode * 397) ^ (int)horizontalAlignment;
                hashCode = (hashCode * 397) ^ wrapText.GetHashCode();
                return hashCode;
            }
        }

        private EnumValue<VerticalAlignmentValues> GetVerticalAlignment()
        {
            EnumValue<VerticalAlignmentValues> result;
            switch (verticalAlignment)
            {
            case ExcelVerticalAlignment.Top:
                result = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Top);
                break;
            case ExcelVerticalAlignment.Center:
                result = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Center);
                break;
            case ExcelVerticalAlignment.Bottom:
                result = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Bottom);
                break;
            case ExcelVerticalAlignment.Default:
                result = null;
                break;
            default:
                throw new InvalidOperationException($"Unknown vertical alignment: {verticalAlignment}");
            }
            return result;
        }

        private EnumValue<HorizontalAlignmentValues> GetHorizontalAlignment()
        {
            EnumValue<HorizontalAlignmentValues> result;
            switch (horizontalAlignment)
            {
            case ExcelHorizontalAlignment.Left:
                result = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Left);
                break;
            case ExcelHorizontalAlignment.Center:
                result = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Center);
                break;
            case ExcelHorizontalAlignment.Right:
                result = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Right);
                break;
            case ExcelHorizontalAlignment.Default:
                result = null;
                break;
            default:
                throw new InvalidOperationException($"Unknown horizontal alignment: {horizontalAlignment}");
            }
            return result;
        }

        private readonly ExcelVerticalAlignment verticalAlignment;
        private readonly ExcelHorizontalAlignment horizontalAlignment;
        private readonly bool wrapText;
    }
}