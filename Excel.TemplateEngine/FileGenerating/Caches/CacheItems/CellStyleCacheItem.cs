using System;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel.TemplateEngine.FileGenerating.Caches.CacheItems
{
    public class CellStyleCacheItem : IEquatable<CellStyleCacheItem>
    {
        public bool Equals(CellStyleCacheItem other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return FillId == other.FillId && FontId == other.FontId && BorderId == other.BorderId && NumberFormatId == other.NumberFormatId && Equals(Alignment, other.Alignment);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((CellStyleCacheItem)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = (int)FillId;
                hashCode = (hashCode * 397) ^ (int)FontId;
                hashCode = (hashCode * 397) ^ (int)BorderId;
                hashCode = (hashCode * 397) ^ (int)NumberFormatId;
                hashCode = (hashCode * 397) ^ (Alignment != null ? Alignment.GetHashCode() : 0);
                return hashCode;
            }
        }

        public CellFormat ToCellFormat()
        {
            return new CellFormat
                {
                    FormatId = 0,
                    FontId = FontId,
                    NumberFormatId = NumberFormatId,
                    FillId = FillId,
                    BorderId = BorderId,
                    Alignment = Alignment?.ToAlignment(),
                    ApplyFill = FillId == 0 ? null : new BooleanValue(true),
                    ApplyBorder = BorderId == 0 ? null : new BooleanValue(true),
                    ApplyNumberFormat = NumberFormatId == 0 ? null : new BooleanValue(true),
                    ApplyAlignment = Alignment == null ? null : new BooleanValue(true),
                    ApplyFont = FontId == 0 ? null : new BooleanValue(true)
                };
        }

        public uint FillId { get; set; }
        public uint FontId { get; set; }
        public uint BorderId { get; set; }
        public uint NumberFormatId { get; set; }
        public AlignmentCacheItem Alignment { get; set; }
    }
}