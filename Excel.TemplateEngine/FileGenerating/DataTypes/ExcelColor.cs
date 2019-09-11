namespace Excel.TemplateEngine.FileGenerating.DataTypes
{
    public class ExcelColor
    {
        public ExcelColor(int alpha, int red, int green, int blue)
        {
            if (red == 0 && green == 0 && blue == 0)
            {
                alpha = 255;
            }
            Red = red;
            Green = green;
            Blue = blue;
            Alpha = alpha;
        }

        public int Red { get; set; }
        public int Green { get; set; }
        public int Blue { get; set; }
        public int Alpha { get; set; }

        public override string ToString() => $"R = {Red}, G = {Green}, B = {Blue}, A = {Alpha}";
    }
}