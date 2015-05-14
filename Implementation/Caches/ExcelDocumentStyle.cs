using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal class ExcelDocumentStyle : IExcelDocumentStyle
    {
        public ExcelDocumentStyle(Stylesheet stylesheet)
        {
            this.stylesheet = stylesheet;
            numberingFormats = new ExcelDocumentNumberingFormats(stylesheet);
            fillStyles = new ExcelDocumentFillStyles(stylesheet);
            bordersStyles = new ExcelDocumentBordersStyles(stylesheet);
            fontStyles = new ExcelDocumentFontStyles(stylesheet);
            cache = new Dictionary<CellStyleCacheItem, uint>();
        }

        public uint AddStyle(ExcelCellStyle style)
        {
            var fillId = fillStyles.AddStyle(style.FillStyle);
            var fontId = fontStyles.AddFont(style.FontStyle);
            var borderId = bordersStyles.AddStyle(style.BordersStyle);
            var numberFormatId = numberingFormats.AddFormat(style.NumberingFormat);
            var alignment = Alignment(style.Alignment);
            var cacheItem = new CellStyleCacheItem
                {
                    FillId = fillId,
                    FontId = fontId,
                    BorderId = borderId,
                    NumberFormatId = numberFormatId,
                    Alignment = alignment
                };
            uint result;
            if(!cache.TryGetValue(cacheItem, out result))
            {
                result = stylesheet.CellFormats.Count;
                stylesheet.CellFormats.Count++;
                stylesheet.CellFormats.AppendChild(cacheItem.ToCellFormat());
                cache.Add(cacheItem, result);
            }
            return result;
        }

        public ExcelCellStyle GetStyle(int styleIndex)
        {
            var cellFormat = (CellFormat)stylesheet
                                             .With(s => s.CellFormats)
                                             .With(cf => cf.ChildElements)
                                             .If(ce => ce.Count > styleIndex)
                                             .Return(ce => ce[styleIndex], null);

            var result = new ExcelCellStyle
                {
                    FillStyle = cellFormat.With(cf => cf.FillId) == null ? null : GetCellFillStyle(cellFormat.FillId.Value),
                    FontStyle = cellFormat.With(cf => cf.FontId) == null ? null : GetCellFontStyle(cellFormat.FontId.Value),
                    NumberingFormat = cellFormat.With(cf => cf.NumberFormatId) == null ? null : GetCellNumberingFormat(cellFormat.NumberFormatId.Value),
                    BordersStyle = cellFormat.With(cf => cf.BorderId) == null ? null : GetCellBordersStyle(cellFormat.BorderId.Value),
                    Alignment = cellFormat.With(cf => cf.Alignment) == null ? null : GetCellAlignment(cellFormat.Alignment)
                };

            return result;
        }

        private ExcelCellAlignment GetCellAlignment(Alignment alignment)
        {
            return new ExcelCellAlignment
                {
                    WrapText = true, //alignment.With(a => a.WrapText) != null && alignment.WrapText.Value, всегда делать перенос по словам
                    HorizontalAlignment = alignment.With(a => a.Horizontal) == null ? ExcelHorizontalAlignment.Default : ToExcelHorizontalAlignment(alignment.Horizontal),
                    VerticalAlignment = alignment.With(a => a.Vertical) == null ? ExcelVerticalAlignment.Default : ToExcelVerticalAlignment(alignment.Vertical)
                };
        }

        private ExcelVerticalAlignment ToExcelVerticalAlignment(EnumValue<VerticalAlignmentValues> vertical)
        {
            switch(vertical.Value)
            {
            case VerticalAlignmentValues.Bottom:
                return ExcelVerticalAlignment.Bottom;
            case VerticalAlignmentValues.Center:
                return ExcelVerticalAlignment.Center;
            case VerticalAlignmentValues.Top:
                return ExcelVerticalAlignment.Top;
            default:
                return ExcelVerticalAlignment.Default;
            }
        }

        private ExcelHorizontalAlignment ToExcelHorizontalAlignment(EnumValue<HorizontalAlignmentValues> horizontal)
        {
            switch(horizontal.Value)
            {
            case HorizontalAlignmentValues.Center:
                return ExcelHorizontalAlignment.Center;
            case HorizontalAlignmentValues.Left:
                return ExcelHorizontalAlignment.Left;
            case HorizontalAlignmentValues.Right:
                return ExcelHorizontalAlignment.Right;
            default:
                return ExcelHorizontalAlignment.Default;
            }
        }

        private ExcelCellBordersStyle GetCellBordersStyle(uint borderId)
        {
            var bordersStyle = (Border)stylesheet
                                           .With(s => s.Borders)
                                           .With(b => b.ChildElements)
                                           .If(ce => ce.Count > borderId)
                                           .Return(ce => ce[(int)borderId], null);
            return new ExcelCellBordersStyle
                {
                    LeftBorder = bordersStyle.With(bs => bs.LeftBorder) == null ? null : GetBorderStyle(bordersStyle.LeftBorder),
                    RightBorder = bordersStyle.With(bs => bs.RightBorder) == null ? null : GetBorderStyle(bordersStyle.RightBorder),
                    TopBorder = bordersStyle.With(bs => bs.TopBorder) == null ? null : GetBorderStyle(bordersStyle.TopBorder),
                    BottomBorder = bordersStyle.With(bs => bs.BottomBorder) == null ? null : GetBorderStyle(bordersStyle.BottomBorder)
                };
        }

        private static ExcelCellBorderStyle GetBorderStyle(BorderPropertiesType border)
        {
            return new ExcelCellBorderStyle
                {
                    Color = border.With(b => b.Color).With(c => c.Rgb) == null ? null : ToExcelColor(border.Color.Rgb),
                    BorderType = border.With(b => b.Style) == null ? ExcelBorderType.None : ToExcelBorderType(border.Style)
                };
        }

        private static ExcelBorderType ToExcelBorderType(EnumValue<BorderStyleValues> borderStyle)
        {
            switch(borderStyle.Value)
            {
            case BorderStyleValues.None:
                return ExcelBorderType.None;
            case BorderStyleValues.Thin:
                return ExcelBorderType.Thin;
            case BorderStyleValues.Medium:
                return ExcelBorderType.Single;
            case BorderStyleValues.Thick:
                return ExcelBorderType.Bold;
            case BorderStyleValues.Double:
                return ExcelBorderType.Double;
            default:
                throw new Exception(string.Format("Unknown border type: {0}", borderStyle));
            }
        }

        private ExcelCellNumberingFormat GetCellNumberingFormat(uint numberFormatId)
        {
            ExcelCellNumberingFormat result;
            if(TryExtractStandartNumberingFormat(numberFormatId, out result))
                return result;

            var numberFormat = (NumberingFormat)stylesheet
                                                    .With(s => s.NumberingFormats)
                                                    .With(nfs => nfs.ChildElements)
                                                    .ReturnEnumerable()
                                                    .FirstOrDefault(ce => ((NumberingFormat)ce)
                                                                              .With(nf => nf.NumberFormatId) != null &&
                                                                          ((NumberingFormat)ce).NumberFormatId.Value == numberFormatId);
            if(numberFormat.With(nf => nf.FormatCode).With(fc => fc.Value) == null)
                return null;

            // ReSharper disable once PossibleNullReferenceException
            return new ExcelCellNumberingFormat(numberFormat.FormatCode.Value);
        }

        private static bool TryExtractStandartNumberingFormat(uint numberingFormat, out ExcelCellNumberingFormat result)
        {
            result = null;
            if(numberingFormat == 2)
            {
                result = new ExcelCellNumberingFormat("0.00");
                return true;
            }

            return false;
        }

        private ExcelCellFontStyle GetCellFontStyle(uint fontId)
        {
            var internalFont = (Font)stylesheet
                                         .With(s => s.Fonts)
                                         .With(f => f.ChildElements)
                                         .If(cf => cf.Count > fontId)
                                         .Return(cf => cf[(int)fontId], null);
            return new ExcelCellFontStyle
                {
                    Bold = internalFont.With(font => font.Bold) != null,
                    Size = internalFont.With(font => font.FontSize) == null ? (int?)null : Convert.ToInt32(internalFont.FontSize.With(fs => fs.Val).Value),
                    Underlined = internalFont.With(font => font.Underline) != null,
                    Color = internalFont.With(font => font.Color).With(c => c.Rgb) == null ? null : ToExcelColor(internalFont.Color.Rgb)
                };
        }

        private ExcelCellFillStyle GetCellFillStyle(uint fillId)
        {
            var internalColor = ((Fill)stylesheet
                                           .With(s => s.Fills)
                                           .With(f => f.ChildElements)
                                           .If(ce => ce.Count > fillId)
                                           .Return(ce => ce[(int)fillId], null))
                .With(f => f.PatternFill)
                .With(pf => pf.ForegroundColor)
                .With(fc => fc.Rgb);

            if(internalColor == null)
                return null;

            var color = ToExcelColor(internalColor);

            return new ExcelCellFillStyle
                {
                    Color = color
                };
        }

        private static ExcelColor ToExcelColor(HexBinaryValue color)
        {
            return new ExcelColor
                {
                    Alpha = Convert.ToInt32(color.Value.Substring(0, 2), 16),
                    Red = Convert.ToInt32(color.Value.Substring(2, 2), 16),
                    Green = Convert.ToInt32(color.Value.Substring(4, 2), 16),
                    Blue = Convert.ToInt32(color.Value.Substring(6, 2), 16),
                };
        }

        private static AlignmentCacheItem Alignment(ExcelCellAlignment cellAlignment)
        {
            if(cellAlignment == null)
                return null;
            if(cellAlignment.HorizontalAlignment == ExcelHorizontalAlignment.Default && cellAlignment.VerticalAlignment == ExcelVerticalAlignment.Default && !cellAlignment.WrapText)
                return null;
            return new AlignmentCacheItem(cellAlignment);
        }

        private readonly Stylesheet stylesheet;
        private readonly ExcelDocumentNumberingFormats numberingFormats;
        private readonly ExcelDocumentFillStyles fillStyles;
        private readonly ExcelDocumentBordersStyles bordersStyles;
        private readonly IExcelDocumentFontStyles fontStyles;
        private readonly IDictionary<CellStyleCacheItem, uint> cache;
    }
}