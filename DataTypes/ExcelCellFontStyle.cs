using System.Collections.Generic;

namespace SKBKontur.Catalogue.ExcelFileGenerator.DataTypes
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
            if(Size != null)
                lines.Add(string.Format("Size = {0}", Size.Value));
            if(Color != null)
                lines.Add(string.Format("Color = {0}", Color));
            if(Underlined)
                lines.Add("Underlined");
            if(Bold)
                lines.Add("Bold");

            if(lines.Count != 0)
                return "\n\t\t\t" + string.Join("\n\t\t\t", lines) + "\n\t\t";

            return "";
        }
    }
}