using System;
using System.Linq;

using DocumentFormat.OpenXml.Spreadsheet;

using Excel.TemplateEngine.FileGenerating.DataTypes;

namespace Excel.TemplateEngine.FileGenerating.Caches.CacheItems
{
    internal class SharedStringCacheItem : IEquatable<SharedStringCacheItem>
    {
        public SharedStringCacheItem(FormattedStringValue value)
        {
            blocks = value.Blocks.Select(block => new FormattedStringValueBlockCacheItem(block)).ToArray();
        }

        public bool Equals(SharedStringCacheItem other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            if (blocks.Length != other.blocks.Length)
                return false;
            return blocks.Zip(other.blocks, (block, otherBlock) => block.Equals(otherBlock)).All(x => x);
        }

        public SharedStringItem ToSharedStringItem()
        {
            return new SharedStringItem(blocks.Select(block => block.ToRun()));
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != GetType()) return false;
            return Equals((SharedStringCacheItem)obj);
        }

        public override int GetHashCode()
        {
            return blocks.Aggregate(0, (hash, block) => hash * 397 ^ block.GetHashCode());
        }

        private readonly FormattedStringValueBlockCacheItem[] blocks;
    }
}