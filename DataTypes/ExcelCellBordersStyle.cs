using System.Collections.Generic;

namespace SKBKontur.Catalogue.ExcelFileGenerator.DataTypes
{
    public class ExcelCellBordersStyle
    {
        public ExcelCellBorderStyle LeftBorder { get; set; }
        public ExcelCellBorderStyle RightBorder { get; set; }
        public ExcelCellBorderStyle TopBorder { get; set; }
        public ExcelCellBorderStyle BottomBorder { get; set; }

        public override string ToString()
        {
            var lines = new List<string>();
            if (LeftBorder != null && LeftBorder.BorderType != ExcelBorderType.None)
                lines.Add(string.Format("LeftBorder = {{{0}}}", LeftBorder));
            if (RightBorder != null && RightBorder.BorderType != ExcelBorderType.None)
                lines.Add(string.Format("RightBorder = {{{0}}}", RightBorder));
            if (TopBorder != null && TopBorder.BorderType != ExcelBorderType.None)
                lines.Add(string.Format("LeftBorder = {{{0}}}", TopBorder));
            if (BottomBorder != null && BottomBorder.BorderType != ExcelBorderType.None)
                lines.Add(string.Format("RightBorder = {{{0}}}", BottomBorder));

            if (lines.Count != 0)
                return "\n\t\t\t" + string.Join("\n\t\t\t", lines) + "\n\t\t";

            return "";
        }
    }
}