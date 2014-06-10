using System;
using System.Linq;

using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems
{
    public class SharedStringCacheItem : IEquatable<SharedStringCacheItem>
    {
        public SharedStringCacheItem(FormattedStringValue value)
        {
            blocks = value.Blocks.Select(block => new FormattedStringValueBlockCacheItem(block)).ToArray();
        }

        public bool Equals(SharedStringCacheItem other)
        {
            if(ReferenceEquals(null, other)) return false;
            if(ReferenceEquals(this, other)) return true;
            if(blocks.Length != other.blocks.Length)
                return false;
            return blocks.Zip(other.blocks, (block, otherBlock) => block.Equals(otherBlock)).All(x => x);
        }

        public SharedStringItem ToSharedStringItem()
        {
            return new SharedStringItem(blocks.Select(block => block.ToRun()));
        }

        public override bool Equals(object obj)
        {
            if(ReferenceEquals(null, obj)) return false;
            if(ReferenceEquals(this, obj)) return true;
            if(obj.GetType() != GetType()) return false;
            return Equals((SharedStringCacheItem)obj);
        }

        public override int GetHashCode()
        {
            return blocks.Aggregate(0, (hash, block) => hash * 397 ^ block.GetHashCode());
        }

        private readonly FormattedStringValueBlockCacheItem[] blocks;
    }

    public class FormattedStringValueBlockCacheItem : IEquatable<FormattedStringValueBlockCacheItem>
    {
        public FormattedStringValueBlockCacheItem(FormattedStringValueBlock formattedStringValueBlock)
        {
            value = formattedStringValueBlock.Value;
            color = formattedStringValueBlock.Color == null ? null : new ColorCacheItem(formattedStringValueBlock.Color);
        }

        public bool Equals(FormattedStringValueBlockCacheItem other)
        {
            if(ReferenceEquals(null, other)) return false;
            if(ReferenceEquals(this, other)) return true;
            return string.Equals(value, other.value) && Equals(color, other.color);
        }

        public override bool Equals(object obj)
        {
            if(ReferenceEquals(null, obj)) return false;
            if(ReferenceEquals(this, obj)) return true;
            if(obj.GetType() != GetType()) return false;
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
                    Text = new Text(value),
                    RunProperties = color == null
                                        ? null
                                        : new RunProperties(color.ToColor())
                };
        }

        private readonly string value;
        private readonly ColorCacheItem color;
    }
}