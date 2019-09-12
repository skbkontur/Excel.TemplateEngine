using System.Collections.Generic;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes
{
    public class ExcelCellBorderStyle
    {
        public ExcelColor Color { get; set; }
        public ExcelBorderType BorderType { get; set; }

        public override string ToString()
        {
            var lines = new List<string>();
            if (Color != null)
                lines.Add($"Color = {Color}");
            lines.Add($"BorderType = {BorderType}");

            return "\n\t\t\t\t" + string.Join("\n\t\t\t\t", lines) + "\n\t\t\t";
        }
    }
}