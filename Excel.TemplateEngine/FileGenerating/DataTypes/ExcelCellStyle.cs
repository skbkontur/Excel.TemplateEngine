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
                lines.Add(string.Format("\tNumberingFormat = {{{0}}}", NumberingFormat));
            if (FillStyle != null && !string.IsNullOrEmpty(FillStyle.ToString()))
                lines.Add(string.Format("\tFillStyle = {{{0}}}", FillStyle));
            if (FontStyle != null && !string.IsNullOrEmpty(FontStyle.ToString()))
                lines.Add(string.Format("\tFontStyle = {{{0}}}", FontStyle));
            if (BordersStyle != null && !string.IsNullOrEmpty(BordersStyle.ToString()))
                lines.Add(string.Format("\tBordersStyle = {{{0}}}", BordersStyle));
            if (Alignment != null && !string.IsNullOrEmpty(Alignment.ToString()))
                lines.Add(string.Format("\tAlignment = {{{0}}}", Alignment));
            return string.Join("\n", lines);
        }
    }
}