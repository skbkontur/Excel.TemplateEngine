using System;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

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
        }

        public void Save()
        {
            stylesheet.Save();
        }

        public uint AddStyle(ExcelCellStyle style)
        {
            var fillId = fillStyles.AddStyle(style.FillStyle);
            var fontId = fontStyles.AddFont(style.FontStyle);
            var borderId = bordersStyles.AddStyle(style.BordersStyle);
            var numberFormatId = numberingFormats.AddFormat(style.NumberingFormat);
            var styleFormatId = stylesheet.CellFormats.Count.Value;
            var alignment = Alignment(style.Alignment);
            stylesheet.CellFormats.Count++;
            stylesheet.CellFormats.AppendChild(new CellFormat
                {
                    FormatId = 0,
                    FontId = fontId,
                    NumberFormatId = numberFormatId,
                    FillId = fillId,
                    BorderId = borderId,
                    Alignment = alignment,
                    ApplyFill = fillId == 0 ? null : new BooleanValue(true),
                    ApplyBorder = borderId == 0 ? null : new BooleanValue(true),
                    ApplyNumberFormat = numberFormatId == 0 ? null : new BooleanValue(true),
                    ApplyAlignment = alignment == null ? null : new BooleanValue(true),
                    ApplyFont = fontId == 0 ? null : new BooleanValue(true)
                });
            return styleFormatId;
        }

        private static Alignment Alignment(ExcelCellAlignment cellAlignment)
        {
            if(cellAlignment == null)
                return null;
            EnumValue<HorizontalAlignmentValues> horizontalAlignment;
            switch(cellAlignment.HorizontalAlignment)
            {
            case ExcelHorizontalAlignment.Left:
                horizontalAlignment = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Left);
                break;
            case ExcelHorizontalAlignment.Center:
                horizontalAlignment = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Center);
                break;
            case ExcelHorizontalAlignment.Right:
                horizontalAlignment = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Right);
                break;
            case ExcelHorizontalAlignment.Default:
                horizontalAlignment = null;
                break;
            default:
                throw new Exception(string.Format("Unknown horizontal alignment: {0}", cellAlignment.HorizontalAlignment));
            }

            EnumValue<VerticalAlignmentValues> verticalAlignment;
            switch(cellAlignment.VerticalAlignment)
            {
            case ExcelVerticalAlignment.Top:
                verticalAlignment = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Top);
                break;
            case ExcelVerticalAlignment.Center:
                verticalAlignment = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Center);
                break;
            case ExcelVerticalAlignment.Bottom:
                verticalAlignment = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Bottom);
                break;
            case ExcelVerticalAlignment.Default:
                verticalAlignment = null;
                break;
            default:
                throw new Exception(string.Format("Unknown vertical alignment: {0}", cellAlignment.VerticalAlignment));
            }
            if(horizontalAlignment == null && verticalAlignment == null && !cellAlignment.WrapText)
                return null;

            return new Alignment
                {
                    Horizontal = horizontalAlignment,
                    Vertical = verticalAlignment,
                    WrapText = cellAlignment.WrapText ? new BooleanValue(true) : null
                };
        }

        private readonly Stylesheet stylesheet;
        private readonly ExcelDocumentNumberingFormats numberingFormats;
        private readonly ExcelDocumentFillStyles fillStyles;
        private readonly ExcelDocumentBordersStyles bordersStyles;
        private readonly IExcelDocumentFontStyles fontStyles;
    }
}