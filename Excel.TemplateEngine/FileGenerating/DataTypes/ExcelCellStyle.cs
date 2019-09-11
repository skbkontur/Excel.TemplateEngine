using System.Collections.Generic;

namespace Excel.TemplateEngine.FileGenerating.DataTypes
{
    public class ExcelCellStyle
    {
        public ExcelCellNumberingFormat NumberingFormat { get; set; }
        public ExcelCellFillStyle FillStyle { get; set; }
        public ExcelCellFontStyle FontStyle { get; set; }
        public ExcelCellBordersStyle BordersStyle { get; set; }
        public ExcelCellAlignment Alignment { get; set; }

        public override string ToString()
        {
            var lines = new List<string>();
            if (NumberingFormat != null && !string.IsNullOrEmpty(NumberingFormat.ToString()))
                lines.Add($"\tNumberingFormat = {{{NumberingFormat}}}");
            if (FillStyle != null && !string.IsNullOrEmpty(FillStyle.ToString()))
                lines.Add($"\tFillStyle = {{{FillStyle}}}");
            if (FontStyle != null && !string.IsNullOrEmpty(FontStyle.ToString()))
                lines.Add($"\tFontStyle = {{{FontStyle}}}");
            if (BordersStyle != null && !string.IsNullOrEmpty(BordersStyle.ToString()))
                lines.Add($"\tBordersStyle = {{{BordersStyle}}}");
            if (Alignment != null && !string.IsNullOrEmpty(Alignment.ToString()))
                lines.Add($"\tAlignment = {{{Alignment}}}");
            return string.Join("\n", lines);
        }
    }
}