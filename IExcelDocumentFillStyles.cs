using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public interface IExcelDocumentFillStyles
    {
        uint AddStyle(ExcelCellFillStyle style);
    }

    internal class ExcelDocumentFillStyles : IExcelDocumentFillStyles
    {
        public ExcelDocumentFillStyles(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            cache = new Dictionary<FillStyleCacheItem, uint>();
        }

        public uint AddStyle(ExcelCellFillStyle format)
        {
            if(format == null)
                return 0;
            var cacheItem = new FillStyleCacheItem(format);
            uint result;
            if(cache.TryGetValue(cacheItem, out result))
                return result;
            if(stylesheet.Fills == null)
            {
                var fills = new Fills {Count = new UInt32Value(0u)};
                if(stylesheet.Fonts != null)
                    stylesheet.InsertAfter(fills, stylesheet.Fonts);
                else if(stylesheet.NumberingFormats != null)
                    stylesheet.InsertAfter(fills, stylesheet.NumberingFormats);
                else
                    stylesheet.InsertAt(fills, 0);
            }
            result = stylesheet.Fills.Count;
            stylesheet.Fills.AppendChild(new Fill
                {
                    PatternFill = new PatternFill(new ForegroundColor {Rgb = cacheItem.GetRGBString()})
                        {
                            PatternType = new EnumValue<PatternValues>(PatternValues.Solid)
                        }
                });
            stylesheet.Fills.Count++;
            return result;
        }

        private readonly Stylesheet stylesheet;

        private readonly Dictionary<FillStyleCacheItem, uint> cache;

        private class FillStyleCacheItem : IEquatable<FillStyleCacheItem>
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
                if(ReferenceEquals(null, obj)) return false;
                if(ReferenceEquals(this, obj)) return true;
                if(obj.GetType() != GetType()) return false;
                return Equals((FillStyleCacheItem)obj);
            }

            public override int GetHashCode()
            {
                return Color.GetHashCode();
            }

            public HexBinaryValue GetRGBString()
            {
                return Color.GetRGBString();
            }

            private ColorCacheItem Color { get; set; }
        }

        private class ColorCacheItem : IEquatable<ColorCacheItem>
        {
            public ColorCacheItem(ExcelColor color)
            {
                Red = color.Red;
                Green = color.Green;
                Blue = color.Blue;
                Alpha = color.Alpha;
            }

            public bool Equals(ColorCacheItem other)
            {
                if(ReferenceEquals(null, other)) return false;
                if(ReferenceEquals(this, other)) return true;
                return Red == other.Red && Green == other.Green && Blue == other.Blue && Alpha == other.Alpha;
            }

            public override bool Equals(object obj)
            {
                if(ReferenceEquals(null, obj)) return false;
                if(ReferenceEquals(this, obj)) return true;
                if(obj.GetType() != GetType()) return false;
                return Equals((ColorCacheItem)obj);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    var hashCode = Red;
                    hashCode = (hashCode * 397) ^ Green;
                    hashCode = (hashCode * 397) ^ Blue;
                    hashCode = (hashCode * 397) ^ Alpha;
                    return hashCode;
                }
            }

            public HexBinaryValue GetRGBString()
            {
                return new HexBinaryValue {Value = string.Format("{0:X2}{1:X2}{2:X2}{3:X2}", Alpha, Red, Green, Blue)};
            }

            private int Red { get; set; }
            private int Green { get; set; }
            private int Blue { get; set; }
            private int Alpha { get; set; }
        }
    }
}