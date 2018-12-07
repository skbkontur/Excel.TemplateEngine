namespace SKBKontur.Catalogue.ExcelFileGenerator.DataTypes
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

        public override string ToString()
        {
            return string.Format("R = {0}, G = {1}, B = {2}, A = {3}", Red, Green, Blue, Alpha);
        }
    }
}