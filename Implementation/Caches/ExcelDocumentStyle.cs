using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Helpers;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.CacheItems;
using SKBKontur.Catalogue.Objects;
using SKBKontur.Catalogue.ServiceLib.Logging;

using ColorType = DocumentFormat.OpenXml.Spreadsheet.ColorType;
using Fill = DocumentFormat.OpenXml.Spreadsheet.Fill;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches
{
    internal class ExcelDocumentStyle : IExcelDocumentStyle
    {
        public ExcelDocumentStyle(Stylesheet stylesheet, Theme theme)
        {
            this.stylesheet = stylesheet;
            numberingFormats = new ExcelDocumentNumberingFormats(stylesheet);
            fillStyles = new ExcelDocumentFillStyles(stylesheet);
            bordersStyles = new ExcelDocumentBordersStyles(stylesheet);
            fontStyles = new ExcelDocumentFontStyles(stylesheet);
            // Not using theme.ThemeElements.ColorScheme.Elements<Color2Type>() here because of wrong order.
            colorSchemeElements = new List<Color2Type>
                {
                    theme.ThemeElements.ColorScheme.Light1Color,
                    theme.ThemeElements.ColorScheme.Dark1Color,
                    theme.ThemeElements.ColorScheme.Light2Color,
                    theme.ThemeElements.ColorScheme.Dark2Color,
                    theme.ThemeElements.ColorScheme.Accent1Color,
                    theme.ThemeElements.ColorScheme.Accent2Color,
                    theme.ThemeElements.ColorScheme.Accent3Color,
                    theme.ThemeElements.ColorScheme.Accent4Color,
                    theme.ThemeElements.ColorScheme.Accent5Color,
                    theme.ThemeElements.ColorScheme.Accent6Color,
                    theme.ThemeElements.ColorScheme.Hyperlink,
                    theme.ThemeElements.ColorScheme.FollowedHyperlinkColor,
                };
            cache = new Dictionary<CellStyleCacheItem, uint>();
            inverseCache = new Dictionary<uint, ExcelCellStyle>();
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
            if (!cache.TryGetValue(cacheItem, out var result))
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
            if (inverseCache.TryGetValue((uint)styleIndex, out var result))
                return result;

            var cellFormat = stylesheet?.CellFormats?.ChildElements?.Count > styleIndex ? (CellFormat)stylesheet.CellFormats.ChildElements[styleIndex] : null;
            result = new ExcelCellStyle
                {
                    FillStyle = cellFormat?.FillId == null ? null : GetCellFillStyle(cellFormat.FillId.Value),
                    FontStyle = cellFormat?.FontId == null ? null : GetCellFontStyle(cellFormat.FontId.Value),
                    NumberingFormat = cellFormat?.NumberFormatId == null ? null : GetCellNumberingFormat(cellFormat.NumberFormatId.Value),
                    BordersStyle = cellFormat?.BorderId == null ? null : GetCellBordersStyle(cellFormat.BorderId.Value),
                    Alignment = cellFormat?.Alignment == null ? null : GetCellAlignment(cellFormat.Alignment)
                };
            inverseCache.Add((uint)styleIndex, result);

            return result;
        }

        private ExcelCellAlignment GetCellAlignment(Alignment alignment)
        {
            return new ExcelCellAlignment
                {
                    WrapText = true, //alignment.With(a => a.WrapText) != null && alignment.WrapText.Value, всегда делать перенос по словам
                    HorizontalAlignment = alignment?.Horizontal == null ? ExcelHorizontalAlignment.Default : ToExcelHorizontalAlignment(alignment.Horizontal),
                    VerticalAlignment = alignment?.Vertical == null ? ExcelVerticalAlignment.Default : ToExcelVerticalAlignment(alignment.Vertical)
                };
        }

        private ExcelVerticalAlignment ToExcelVerticalAlignment(EnumValue<VerticalAlignmentValues> vertical)
        {
            switch (vertical.Value)
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
            switch (horizontal.Value)
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
            var bordersStyle = stylesheet?.Borders?.ChildElements?.Count > borderId ? (Border)stylesheet.Borders.ChildElements[(int)borderId] : null;
            return new ExcelCellBordersStyle
                {
                    LeftBorder = bordersStyle?.LeftBorder == null ? null : GetBorderStyle(bordersStyle.LeftBorder),
                    RightBorder = bordersStyle?.RightBorder == null ? null : GetBorderStyle(bordersStyle.RightBorder),
                    TopBorder = bordersStyle?.TopBorder == null ? null : GetBorderStyle(bordersStyle.TopBorder),
                    BottomBorder = bordersStyle?.BottomBorder == null ? null : GetBorderStyle(bordersStyle.BottomBorder)
                };
        }

        private ExcelCellBorderStyle GetBorderStyle(BorderPropertiesType border)
        {
            return new ExcelCellBorderStyle
                {
                    Color = ToExcelColor(border?.Color),
                    BorderType = border?.Style == null ? ExcelBorderType.None : ToExcelBorderType(border.Style)
                };
        }

        private static ExcelBorderType ToExcelBorderType(EnumValue<BorderStyleValues> borderStyle)
        {
            switch (borderStyle.Value)
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
                throw new InvalidProgramStateException($"Unknown border type: {borderStyle}");
            }
        }

        private ExcelCellNumberingFormat GetCellNumberingFormat(uint numberFormatId)
        {
            if (TryExtractStandartNumberingFormat(numberFormatId, out var result))
                return result;

            var numberFormat = (NumberingFormat)stylesheet?.NumberingFormats?.ChildElements
                                                          ?.FirstOrDefault(ce => ((NumberingFormat)ce)?.NumberFormatId != null &&
                                                                                 ((NumberingFormat)ce).NumberFormatId.Value == numberFormatId);
            if (numberFormat?.FormatCode?.Value == null)
                return null;

            // ReSharper disable once PossibleNullReferenceException
            return new ExcelCellNumberingFormat(numberFormat.FormatCode.Value);
        }

        private static bool TryExtractStandartNumberingFormat(uint numberingFormat, out ExcelCellNumberingFormat result)
        {
            result = null;
            if (numberingFormat == 2)
            {
                result = new ExcelCellNumberingFormat("0.00");
                return true;
            }

            return false;
        }

        private ExcelCellFontStyle GetCellFontStyle(uint fontId)
        {
            var internalFont = stylesheet?.Fonts?.ChildElements?.Count > fontId ? (Font)stylesheet.Fonts.ChildElements[(int)fontId] : null;
            return new ExcelCellFontStyle
                {
                    Bold = internalFont?.Bold != null,
                    Size = internalFont?.FontSize == null ? (int?)null : Convert.ToInt32(internalFont.FontSize?.Val.Value),
                    Underlined = internalFont?.Underline != null,
                    Color = ToExcelColor(internalFont?.Color)
                };
        }

        private ExcelCellFillStyle GetCellFillStyle(uint fillId)
        {
            var fill = stylesheet?.Fills?.ChildElements?.Count > fillId ? (Fill)stylesheet.Fills.ChildElements[(int)fillId] : null;
            var color = ToExcelColor(fill?.PatternFill?.ForegroundColor);

            if (color == null)
                return null;

            return new ExcelCellFillStyle {Color = color};
        }

        [CanBeNull]
        private ExcelColor ToExcelColor([CanBeNull] ColorType color)
        {
            if (color == null)
                return null;
            if (color.Rgb?.HasValue == true)
            {
                return RgbStringToExcelColor(color.Rgb.Value);
            }
            if (color.Theme?.HasValue == true)
            {
                var theme = color.Theme.Value;
                var tint = color.Tint?.Value ?? 0;
                return ThemeToExcelColor(theme, tint);
            }
            return null;
        }

        [CanBeNull]
        private ExcelColor RgbStringToExcelColor([NotNull] string hexRgbColor)
        {
            if (hexRgbColor.Length == 6)
                hexRgbColor = "FF" + hexRgbColor;
            return new ExcelColor
                {
                    Alpha = Convert.ToInt32(hexRgbColor.Substring(0, 2), 16),
                    Red = Convert.ToInt32(hexRgbColor.Substring(2, 2), 16),
                    Green = Convert.ToInt32(hexRgbColor.Substring(4, 2), 16),
                    Blue = Convert.ToInt32(hexRgbColor.Substring(6, 2), 16),
                };
        }

        [CanBeNull]
        private ExcelColor ThemeToExcelColor(uint theme, double tint)
        {
            if (theme >= colorSchemeElements.Count)
                throw new InvalidProgramStateException($"Theme with id '{theme}' not found");
            var color2Type = colorSchemeElements[(int)theme];
            var rgbColor = color2Type?.RgbColorModelHex?.Val?.Value ?? color2Type?.SystemColor?.LastColor?.Value;
            if (rgbColor == null)
            {
                Log.For(this).Error("Failed to get rgbColor from theme");
                return null;
            }
            var hls = ColorConverter.RgbToHls(RgbStringToExcelColor(rgbColor));
            if (tint < 0)
                hls.L = hls.L * (1.0 + tint);
            else
                hls.L = hls.L * (1.0 - tint) + tint;
            return ColorConverter.HslToRgb(hls);
        }

        private static AlignmentCacheItem Alignment(ExcelCellAlignment cellAlignment)
        {
            if (cellAlignment == null)
                return null;
            if (cellAlignment.HorizontalAlignment == ExcelHorizontalAlignment.Default && cellAlignment.VerticalAlignment == ExcelVerticalAlignment.Default && !cellAlignment.WrapText)
                return null;
            return new AlignmentCacheItem(cellAlignment);
        }

        private readonly Stylesheet stylesheet;
        private readonly ExcelDocumentNumberingFormats numberingFormats;
        private readonly ExcelDocumentFillStyles fillStyles;
        private readonly ExcelDocumentBordersStyles bordersStyles;
        private readonly IExcelDocumentFontStyles fontStyles;
        private readonly List<Color2Type> colorSchemeElements;
        private readonly IDictionary<CellStyleCacheItem, uint> cache;
        private readonly IDictionary<uint, ExcelCellStyle> inverseCache;
    }
}