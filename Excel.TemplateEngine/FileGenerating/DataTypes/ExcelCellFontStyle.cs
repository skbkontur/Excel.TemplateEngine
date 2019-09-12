using System.Collections.Generic;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes
{
    public class ExcelCellFontStyle
    {
        public int? Size { get; set; }
        public ExcelColor Color { get; set; }
        public bool Underlined { get; set; }
        public bool Bold { get; set; }

        public override string ToString()
        {
            var lines = new List<string>();
            if (Size != null)
                lines.Add($"Size = {Size.Value}");
            if (Color != null)
                lines.Add($"Color = {Color}");
            if (Underlined)
                lines.Add("Underlined");
            if (Bold)
                lines.Add("Bold");

            if (lines.Count != 0)
                return "\n\t\t\t" + string.Join("\n\t\t\t", lines) + "\n\t\t";

            return "";
        }
    }
}