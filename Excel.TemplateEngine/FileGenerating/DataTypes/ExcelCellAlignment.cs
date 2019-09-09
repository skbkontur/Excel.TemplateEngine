using System.Collections.Generic;

namespace Excel.TemplateEngine.FileGenerating.DataTypes
{
    public class ExcelCellAlignment
    {
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }
        public ExcelVerticalAlignment VerticalAlignment { get; set; }
        public bool WrapText { get; set; }

        public override string ToString()
        {
            var lines = new List<string>
                {
                    $"HorizontalAlignment = {HorizontalAlignment}",
                    $"VerticalAlignment = {VerticalAlignment}"
                };

            if (WrapText)
                lines.Add("WrapText");

            return "\n\t\t\t" + string.Join("\n\t\t\t", lines) + "\n\t\t";
        }
    }
}