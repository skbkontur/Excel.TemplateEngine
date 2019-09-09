using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Helpers
{
    public static class ExcelCellStyleHelpers
    {
        public static ExcelCellStyle FontSize(this ExcelCellStyle style, int fontSize)
        {
            var result = serializer.Copy(style);
            if (result.FontStyle == null)
                result.FontStyle = new ExcelCellFontStyle();
            result.FontStyle.Size = fontSize;
            return result;
        }

        public static ExcelCellStyle FontColor(this ExcelCellStyle style, ExcelColor color)
        {
            var result = serializer.Copy(style);
            if (result.FontStyle == null)
                result.FontStyle = new ExcelCellFontStyle();
            result.FontStyle.Color = color;
            return result;
        }

        public static ExcelCellStyle FillColor(this ExcelCellStyle style, ExcelColor color)
        {
            var result = serializer.Copy(style);
            if (result.FillStyle == null)
                result.FillStyle = new ExcelCellFillStyle();
            result.FillStyle.Color = color;
            return result;
        }

        public static ExcelCellStyle Bold(this ExcelCellStyle style)
        {
            var result = serializer.Copy(style);
            if (result.FontStyle == null)
                result.FontStyle = new ExcelCellFontStyle();
            result.FontStyle.Bold = true;
            return result;
        }

        public static ExcelCellStyle Underlined(this ExcelCellStyle style)
        {
            var result = serializer.Copy(style);
            if (result.FontStyle == null)
                result.FontStyle = new ExcelCellFontStyle();
            result.FontStyle.Underlined = true;
            return result;
        }

        public static ExcelCellStyle Borders(this ExcelCellStyle style, ExcelCellBordersStyle borders)
        {
            var result = serializer.Copy(style);
            result.BordersStyle = borders;
            return result;
        }

        public static ExcelCellStyle Numeric(this ExcelCellStyle style, int precision)
        {
            var result = serializer.Copy(style);
            if (result.NumberingFormat == null)
                result.NumberingFormat = new ExcelCellNumberingFormat(precision);
            return result;
        }

        public static ExcelCellStyle LeftBorder(this ExcelCellStyle style, ExcelBorderType borderType = ExcelBorderType.Thin, ExcelColor color = null)
        {
            color = color ?? ExcelColors.Black;
            var result = serializer.Copy(style);
            if (result.BordersStyle == null)
                result.BordersStyle = new ExcelCellBordersStyle();
            result.BordersStyle.LeftBorder = new ExcelCellBorderStyle
                {
                    BorderType = borderType,
                    Color = color
                };
            return result;
        }

        public static ExcelCellStyle RightBorder(this ExcelCellStyle style, ExcelBorderType borderType = ExcelBorderType.Thin, ExcelColor color = null)
        {
            color = color ?? ExcelColors.Black;
            var result = serializer.Copy(style);
            if (result.BordersStyle == null)
                result.BordersStyle = new ExcelCellBordersStyle();
            result.BordersStyle.RightBorder = new ExcelCellBorderStyle
                {
                    BorderType = borderType,
                    Color = color
                };
            return result;
        }

        public static ExcelCellStyle TopBorder(this ExcelCellStyle style, ExcelBorderType borderType = ExcelBorderType.Thin, ExcelColor color = null)
        {
            color = color ?? ExcelColors.Black;
            var result = serializer.Copy(style);
            if (result.BordersStyle == null)
                result.BordersStyle = new ExcelCellBordersStyle();
            result.BordersStyle.TopBorder = new ExcelCellBorderStyle
                {
                    BorderType = borderType,
                    Color = color
                };
            return result;
        }

        public static ExcelCellStyle BottomBorder(this ExcelCellStyle style, ExcelBorderType borderType = ExcelBorderType.Thin, ExcelColor color = null)
        {
            color = color ?? ExcelColors.Black;
            var result = serializer.Copy(style);
            if (result.BordersStyle == null)
                result.BordersStyle = new ExcelCellBordersStyle();
            result.BordersStyle.BottomBorder = new ExcelCellBorderStyle
                {
                    BorderType = borderType,
                    Color = color
                };
            return result;
        }

        public static ExcelCellStyle WrapText(this ExcelCellStyle style)
        {
            var result = serializer.Copy(style);
            if (result.Alignment == null)
                result.Alignment = new ExcelCellAlignment();
            result.Alignment.WrapText = true;
            return result;
        }

        public static ExcelCellStyle Alignment(this ExcelCellStyle style, ExcelVerticalAlignment alignment)
        {
            var result = serializer.Copy(style);
            if (result.Alignment == null)
                result.Alignment = new ExcelCellAlignment();
            result.Alignment.VerticalAlignment = alignment;
            return result;
        }

        public static ExcelCellStyle Alignment(this ExcelCellStyle style, ExcelHorizontalAlignment alignment)
        {
            var result = serializer.Copy(style);
            if (result.Alignment == null)
                result.Alignment = new ExcelCellAlignment();
            result.Alignment.HorizontalAlignment = alignment;
            return result;
        }

        private static readonly ISerializer serializer = GrobufSerializers.AllFieldsSerializer;
    }
}