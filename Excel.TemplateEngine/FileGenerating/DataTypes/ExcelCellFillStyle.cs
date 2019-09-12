namespace SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes
{
    public class ExcelCellFillStyle
    {
        public ExcelColor Color { get; set; }

        public override string ToString() => $"FillColor = {{{Color}}}";
    }
}