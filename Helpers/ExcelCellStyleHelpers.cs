using GroBuf;
using GroBuf.DataMembersExtracters;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Helpers
{
    public static class ExcelCellStyleHelpers
    {
        public static ExcelCellStyle FontSize(this ExcelCellStyle style, int fontSize)
        {
            var result = serializer.Copy(style);
            if(result.FontStyle == null)
                result.FontStyle = new ExcelCellFontStyle();
            result.FontStyle.Size = fontSize;
            return result;
        }

        public static ExcelCellStyle FontColor(this ExcelCellStyle style, ExcelColor color)
        {
            var result = serializer.Copy(style);
            if(result.FontStyle == null)
                result.FontStyle = new ExcelCellFontStyle();
            result.FontStyle.Color = color;
            return result;
        }

        public static ExcelCellStyle Bold(this ExcelCellStyle style)
        {
            var result = serializer.Copy(style);
            if(result.FontStyle == null)
                result.FontStyle = new ExcelCellFontStyle();
            result.FontStyle.Bold = true;
            return result;
        }

        public static ExcelCellStyle Underlined(this ExcelCellStyle style)
        {
            var result = serializer.Copy(style);
            if(result.FontStyle == null)
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
            if(result.NumberingFormat == null)
                result.NumberingFormat = new ExcelCellNumberingFormat();
            result.NumberingFormat.Precision = precision;
            return result;
        }

        private static readonly ISerializer serializer = new Serializer(new AllFieldsExtractor());
    }
}