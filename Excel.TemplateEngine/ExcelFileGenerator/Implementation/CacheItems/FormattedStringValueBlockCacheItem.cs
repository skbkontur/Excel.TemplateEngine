using System;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems
{
    public class FormattedStringValueBlockCacheItem : IEquatable<FormattedStringValueBlockCacheItem>
    {
        public FormattedStringValueBlockCacheItem(FormattedStringValueBlock formattedStringValueBlock)
        {
            value = formattedStringValueBlock.Value;
            color = formattedStringValueBlock.Color == null ? null : new ColorCacheItem(formattedStringValueBlock.Color);
        }

        public bool Equals(FormattedStringValueBlockCacheItem other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(value, other.value) && Equals(color, other.color);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((FormattedStringValueBlockCacheItem)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((value != null ? value.GetHashCode() : 0) * 397) ^ (color != null ? color.GetHashCode() : 0);
            }
        }

        public Run ToRun()
        {
            return new Run
                {
                    Text = new Text(value) {Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve)},
                    RunProperties = color == null
                                        ? null
                                        : new RunProperties(color.ToColor<Color>())
                };
        }

        private readonly string value;
        private readonly ColorCacheItem color;
    }
}